/**
Thin wrapper around the optional `dbus-next` dependency.

The session bus is the only cross-desktop channel Wayland leaves us for
window information: GNOME Shell extensions and KWin scripts both speak it.
`dbus-next` is pure JavaScript (its `usocket` addon is optional), but it is
still loaded lazily so that a missing or broken install degrades to the
X11 fallback instead of breaking `activeWindow()` outright.
*/

let modulePromise;
let sessionBus;
const proxyCache = new Map();

async function importDbus() {
	try {
		const module = await import('dbus-next');
		return module.default ?? module;
	} catch {
		return undefined;
	}
}

export async function loadDbus() {
	modulePromise ||= importDbus();
	return modulePromise;
}

export async function getSessionBus() {
	const dbus = await loadDbus();
	if (!dbus) {
		return undefined;
	}

	if (!sessionBus) {
		const bus = dbus.sessionBus();
		// Without a listener a connection failure is an uncaught 'error' event.
		bus.on('error', () => {
			if (sessionBus === bus) {
				sessionBus = undefined;
				proxyCache.clear();
			}
		});
		sessionBus = bus;
	}

	return sessionBus;
}

/**
`ProxyObject#getInterface` returns `undefined` for an interface the object
does not implement, and GNOME Shell answers introspection of an unknown
object path with an empty node rather than an error. Without this check a
missing extension surfaced as a TypeError on the method call, which looked
like a transient failure instead of "provider not installed".
*/
export function requireInterface(object, interfaceName, {name, path} = {}) {
	const iface = object?.getInterface(interfaceName);
	if (!iface) {
		const error = new Error(`Interface ${interfaceName} is not available at ${path} on ${name}`);
		error.type = 'org.freedesktop.DBus.Error.UnknownInterface';
		throw error;
	}

	return iface;
}

async function createInterface(bus, name, path, interfaceName) {
	const object = await bus.getProxyObject(name, path);
	return requireInterface(object, interfaceName, {name, path});
}

export async function getInterface(name, path, interfaceName) {
	const bus = await getSessionBus();
	if (!bus) {
		throw new Error('dbus-next is not available');
	}

	const key = `${name}\0${path}\0${interfaceName}`;
	if (!proxyCache.has(key)) {
		// Introspection happens once per proxy; cache it so 1s polling stays cheap.
		proxyCache.set(key, createInterface(bus, name, path, interfaceName));
	}

	try {
		return await proxyCache.get(key);
	} catch (error) {
		proxyCache.delete(key);
		throw error;
	}
}

// eslint-disable-next-line max-params
export async function callMethod(name, path, interfaceName, method, ...methodArguments) {
	const iface = await getInterface(name, path, interfaceName);
	try {
		return await iface[method](...methodArguments);
	} catch (error) {
		if (isServiceGone(error)) {
			proxyCache.delete(`${name}\0${path}\0${interfaceName}`);
		}

		throw error;
	}
}

/**
The bus's unique connection name (`:1.42`), available once connected.
KWin scripts call us back on it, exactly the way `kdotool` works.
*/
export async function getUniqueName() {
	const bus = await getSessionBus();
	if (!bus) {
		throw new Error('dbus-next is not available');
	}

	if (!bus.name) {
		// Any round-trip completes the Hello handshake that assigns the name.
		await bus.getProxyObject('org.freedesktop.DBus', '/org/freedesktop/DBus');
	}

	return bus.name;
}

const GONE_ERRORS = new Set([
	'org.freedesktop.DBus.Error.ServiceUnknown',
	'org.freedesktop.DBus.Error.NameHasNoOwner',
	'org.freedesktop.DBus.Error.UnknownObject',
	'org.freedesktop.DBus.Error.UnknownInterface',
	'org.freedesktop.DBus.Error.UnknownMethod',
	'org.freedesktop.DBus.Error.NoReply',
	'org.freedesktop.DBus.Error.Disconnected',
]);

/**
Whether an error means "this provider is not installed / not running", as
opposed to a transient failure for a provider that does exist.
*/
export function isServiceGone(error) {
	const type = error?.type || error?.name || '';
	return GONE_ERRORS.has(type) || /dbus-next is not available|DBUS_SESSION_BUS_ADDRESS|ENOENT|ECONNREFUSED/.test(String(error?.message));
}
