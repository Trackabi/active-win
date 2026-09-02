import {callMethod, isServiceGone} from './dbus.js';
import {buildResult, enrichWithProcess} from './process-info.js';

/**
GNOME Shell backends.

Mutter exposes no focused-window API, so GNOME needs a Shell extension.
Trackabi Timer bundles its own (`active-window@trackabi.com`, interface
`io.trackabi.ActiveWindow`) and installs it automatically; three popular
third-party ones publish the same information, and we support them too so a
user who already has any of them installed is covered:

- Focused Window D-Bus (extensions.gnome.org/extension/5592) — what
  ActivityWatch's `awatcher` uses.
- Window Calls Extended (extensions.gnome.org/extension/4974).
- Window Calls (extensions.gnome.org/extension/4724).

Installing an extension on Wayland requires logging out and back in.
*/

const SHELL = 'org.gnome.Shell';

export const SETUP_HINT = 'GNOME needs a Shell extension to report the focused window. Log out and back in to activate the bundled one, or install "Focused Window D-Bus" (https://extensions.gnome.org/extension/5592/focused-window-d-bus/).';

const parseJson = value => {
	if (typeof value !== 'string' || value === '') {
		return undefined;
	}

	try {
		return JSON.parse(value);
	} catch {
		return undefined;
	}
};

/**
Extension methods raise a D-Bus error (not a service-gone one) when nothing
is focused, e.g. on an empty desktop. Treat that as "no window".
*/
const noWindowOnError = async promise => {
	try {
		return await promise;
	} catch (error) {
		if (isServiceGone(error)) {
			throw error;
		}

		return undefined;
	}
};

const fromFields = (backend, {title, id, appId, pid, bounds}) => enrichWithProcess(buildResult({
	backend,
	title,
	id,
	appId,
	name: appId,
	processId: pid,
	bounds,
}));

const fromExtensionWindow = (backend, window) => {
	if (!window || typeof window !== 'object') {
		return undefined;
	}

	return fromFields(backend, {
		title: window.title,
		id: window.id,
		appId: window.wm_class || window.wm_class_instance,
		pid: window.pid,
		bounds: window,
	});
};

export const trackabiActiveWindow = {
	name: 'gnome-trackabi',
	async activeWindow() {
		const json = await noWindowOnError(callMethod(
			SHELL,
			'/io/trackabi/ActiveWindow',
			'io.trackabi.ActiveWindow',
			'Get',
		));
		return fromExtensionWindow(this.name, parseJson(json));
	},
};

export const focusedWindowDbus = {
	name: 'gnome-focused-window-dbus',
	async activeWindow() {
		const json = await noWindowOnError(callMethod(
			SHELL,
			'/org/gnome/shell/extensions/FocusedWindow',
			'org.gnome.shell.extensions.FocusedWindow',
			'Get',
		));
		return fromExtensionWindow(this.name, parseJson(json));
	},
};

export const windowCallsExtended = {
	name: 'gnome-window-calls-extended',
	async activeWindow() {
		const path = '/org/gnome/Shell/Extensions/WindowsExt';
		const iface = 'org.gnome.Shell.Extensions.WindowsExt';
		const [title, pid, wmClass] = await Promise.all([
			noWindowOnError(callMethod(SHELL, path, iface, 'FocusTitle')),
			noWindowOnError(callMethod(SHELL, path, iface, 'FocusPID')),
			noWindowOnError(callMethod(SHELL, path, iface, 'FocusClass')),
		]);
		if (!wmClass && !title) {
			return undefined;
		}

		return fromFields(this.name, {title, pid, appId: wmClass});
	},
};

export const windowCalls = {
	name: 'gnome-window-calls',
	async activeWindow() {
		const path = '/org/gnome/Shell/Extensions/Windows';
		const iface = 'org.gnome.Shell.Extensions.Windows';
		const list = parseJson(await callMethod(SHELL, path, iface, 'List'));
		const focused = Array.isArray(list) ? list.find(window => window.focus) : undefined;
		if (!focused) {
			return undefined;
		}

		const title = await noWindowOnError(callMethod(SHELL, path, iface, 'GetTitle', focused.id));
		return fromExtensionWindow(this.name, {...focused, title});
	},
};

export const backends = [trackabiActiveWindow, focusedWindowDbus, windowCallsExtended, windowCalls];
