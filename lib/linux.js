import {detectSession} from './linux/session.js';
import * as x11 from './linux/x11.js';
import * as hyprland from './linux/hyprland.js';
import * as sway from './linux/sway.js';
import * as kwin from './linux/kwin.js';
import {backends as gnomeBackends, SETUP_HINT as GNOME_SETUP_HINT} from './linux/gnome.js';
import {isServiceGone} from './linux/dbus.js';

/**
Linux entry point.

X11 has one answer (`xprop`). Wayland has none that works everywhere, so we
keep an ordered list of compositor-specific backends and use the first one
that responds. A backend that throws is considered unavailable for a while
(missing tool, extension not installed) and skipped; one that resolves with
`undefined` is working but reports that no window is focused.

The X11 backend is always the last candidate: under XWayland it still
identifies X11 apps, which is better than nothing.
*/

const UNAVAILABLE_TTL = 30_000;

const state = {
	backend: undefined,
	unavailableUntil: new Map(),
	lastError: new Map(),
};

const asyncBackend = (name, module) => ({
	name,
	activeWindow: module.activeWindow,
	activeWindowSync: module.activeWindowSync,
	openWindows: module.openWindows,
	openWindowsSync: module.openWindowsSync,
});

const X11_BACKEND = asyncBackend('x11', x11);
const HYPRLAND_BACKEND = asyncBackend('hyprland', hyprland);
const SWAY_BACKEND = asyncBackend('sway', sway);
const KWIN_BACKEND = {name: 'kwin', activeWindow: kwin.activeWindow};

const isWaylandCapable = backend => backend !== X11_BACKEND;

export function candidateBackends(session = detectSession()) {
	if (!session.isWayland) {
		return [X11_BACKEND];
	}

	const candidates = [];
	if (session.isHyprland) {
		candidates.push(HYPRLAND_BACKEND);
	}

	if (session.isSway) {
		candidates.push(SWAY_BACKEND);
	}

	if (session.isGnome) {
		candidates.push(...gnomeBackends);
	}

	if (session.isKde) {
		candidates.push(KWIN_BACKEND);
	}

	if (candidates.length === 0) {
		// Unknown compositor: probe everything, cheapest first.
		candidates.push(HYPRLAND_BACKEND, SWAY_BACKEND, ...gnomeBackends, KWIN_BACKEND);
	}

	if (session.hasX11Display) {
		candidates.push(X11_BACKEND);
	}

	return candidates;
}

const isUnavailable = backend => (state.unavailableUntil.get(backend.name) ?? 0) > Date.now();

const markUnavailable = (backend, error) => {
	state.unavailableUntil.set(backend.name, Date.now() + UNAVAILABLE_TTL);
	state.lastError.set(backend.name, error);
	if (state.backend === backend) {
		state.backend = undefined;
	}
};

const markAvailable = backend => {
	state.unavailableUntil.delete(backend.name);
	state.lastError.delete(backend.name);
	state.backend = backend;
};

async function runFirst(method, session) {
	for (const backend of candidateBackends(session)) {
		if (typeof backend[method] !== 'function' || isUnavailable(backend)) {
			continue;
		}

		try {
			// eslint-disable-next-line no-await-in-loop
			const result = await backend[method]();
			markAvailable(backend);
			return result;
		} catch (error) {
			markUnavailable(backend, error);
		}
	}

	return undefined;
}

function runFirstSync(method, session) {
	for (const backend of candidateBackends(session)) {
		if (typeof backend[method] !== 'function' || isUnavailable(backend)) {
			continue;
		}

		try {
			const result = backend[method]();
			markAvailable(backend);
			return result;
		} catch (error) {
			markUnavailable(backend, error);
		}
	}

	return undefined;
}

export async function activeWindow() {
	try {
		return await runFirst('activeWindow');
	} catch {
		return undefined;
	}
}

/**
Sync variant: D-Bus backends (GNOME, KWin) are asynchronous by nature, so
this only covers Hyprland, Sway and X11.
*/
export function activeWindowSync() {
	try {
		return runFirstSync('activeWindowSync');
	} catch {
		return undefined;
	}
}

export async function openWindows() {
	try {
		return await runFirst('openWindows');
	} catch {
		return undefined;
	}
}

export function openWindowsSync() {
	try {
		return runFirstSync('openWindowsSync');
	} catch {
		return undefined;
	}
}

const describeError = error => {
	if (!error) {
		return undefined;
	}

	const type = error.type || error.code;
	return type ? `${type}: ${error.message}` : String(error.message ?? error);
};

/**
Explain what tracking can and cannot do in the current session, so an app
can warn the user (e.g. "install this GNOME extension") instead of silently
logging nothing.

Runs one probe when no backend has been selected yet.
*/
export async function linuxTrackingStatus() {
	const session = detectSession();
	if (!state.backend || isUnavailable(state.backend)) {
		await activeWindow();
	}

	const {backend} = state;
	const candidates = candidateBackends(session);
	const waylandCapable = !session.isWayland || Boolean(backend && isWaylandCapable(backend));

	let hint;
	if (session.isWayland && !waylandCapable) {
		if (session.isGnome) {
			hint = GNOME_SETUP_HINT;
		} else if (session.isKde) {
			hint = 'KWin scripting is unavailable; check that KWin is running and D-Bus is reachable.';
		} else if (candidates.length <= 1) {
			hint = 'This Wayland compositor is not supported; only X11 (XWayland) apps can be identified.';
		} else {
			hint = 'No supported Wayland window provider responded; only X11 (XWayland) apps can be identified.';
		}
	}

	return {
		platform: 'linux',
		sessionType: session.sessionType,
		isWayland: session.isWayland,
		desktop: session.desktop,
		backend: backend?.name,
		waylandCapable,
		hint,
		candidates: candidates.map(candidate => ({
			name: candidate.name,
			unavailable: isUnavailable(candidate),
			error: describeError(state.lastError.get(candidate.name)),
			serviceGone: state.lastError.has(candidate.name) ? isServiceGone(state.lastError.get(candidate.name)) : undefined,
		})),
	};
}

export {detectSession} from './linux/session.js';
