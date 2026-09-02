import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import process from 'node:process';
import {promisify} from 'node:util';
import {
	callMethod,
	getSessionBus,
	getUniqueName,
	loadDbus,
} from './dbus.js';
import {buildResult, enrichWithProcess} from './process-info.js';

/**
KDE Plasma (KWin) backend.

KWin has no "active window" D-Bus method, but it will load and run a script
for us over `org.kde.kwin.Scripting`. The script watches focus and caption
changes and pushes them back to *our* unique bus name — the same trick
`kdotool` uses, except our script stays resident so `activeWindow()` is a
cache read instead of a script load per call.

Works on Plasma 5 (`workspace.activeClient`) and Plasma 6
(`workspace.activeWindow`), X11 and Wayland alike.
*/

const writeFile = promisify(fs.writeFile);

export const name = 'kwin';

const KWIN = 'org.kde.KWin';
const SCRIPTING_PATH = '/Scripting';
const SCRIPTING_INTERFACE = 'org.kde.kwin.Scripting';
const SCRIPT_INTERFACE = 'org.kde.kwin.Script';
const PLUGIN_NAME = 'get-windows-active-window';

const RECEIVER_PATH = '/io/trackabi/GetWindows';
const RECEIVER_INTERFACE = 'io.trackabi.GetWindows';

const FIRST_REPORT_TIMEOUT = 1500;

const buildScript = uniqueName => `
	const REPORT_TARGET = ${JSON.stringify(uniqueName)};
	let current = null;

	const currentWindow = () => (workspace.activeWindow !== undefined ? workspace.activeWindow : workspace.activeClient);

	const report = () => {
		const w = currentWindow();
		let payload = null;
		if (w) {
			const g = w.frameGeometry || w.geometry || {};
			payload = {
				caption: w.caption,
				resourceClass: w.resourceClass,
				resourceName: w.resourceName,
				desktopFileName: w.desktopFileName,
				pid: w.pid,
				internalId: String(w.internalId),
				x: g.x, y: g.y, width: g.width, height: g.height,
			};
		}
		callDBus(REPORT_TARGET, ${JSON.stringify(RECEIVER_PATH)}, ${JSON.stringify(RECEIVER_INTERFACE)}, 'Report', JSON.stringify(payload));
	};

	const track = (w) => {
		if (current && current.captionChanged) {
			try { current.captionChanged.disconnect(report); } catch (e) {}
		}
		current = w;
		if (w && w.captionChanged) {
			w.captionChanged.connect(report);
		}
		report();
	};

	if (workspace.windowActivated) {
		workspace.windowActivated.connect(track);
	} else if (workspace.clientActivated) {
		workspace.clientActivated.connect(track);
	}
	track(currentWindow());
`;

class KWinBridge {
	constructor() {
		this.ready = false;
		this.loading = undefined;
		this.last = undefined;
		this.lastReportAt = 0;
		this.waiters = [];
	}

	async ensure() {
		if (this.ready) {
			return;
		}

		this.loading ||= this.load();
		try {
			await this.loading;
		} finally {
			this.loading = undefined;
		}
	}

	async load() {
		const dbus = await loadDbus();
		const bus = await getSessionBus();
		if (!dbus || !bus) {
			throw new Error('dbus-next is not available');
		}

		if (!this.receiver) {
			const {Interface} = dbus.interface;
			const onReport = json => this.onReport(json);
			class Receiver extends Interface {
				Report(json) {
					onReport(json);
				}
			}
			Receiver.configureMembers({methods: {Report: {inSignature: 's', outSignature: ''}}});
			this.receiver = new Receiver(RECEIVER_INTERFACE);
			bus.export(RECEIVER_PATH, this.receiver);
			bus.on('error', () => {
				this.ready = false;
				this.receiver = undefined;
			});
		}

		const uniqueName = await getUniqueName();
		const scriptPath = path.join(os.tmpdir(), `${PLUGIN_NAME}-${process.pid}.js`);
		await writeFile(scriptPath, buildScript(uniqueName));

		// A resident script from a previous (crashed) run reports to a dead name; replace it.
		try {
			await callMethod(KWIN, SCRIPTING_PATH, SCRIPTING_INTERFACE, 'unloadScript', PLUGIN_NAME);
		} catch {}

		const scriptId = await callMethod(KWIN, SCRIPTING_PATH, SCRIPTING_INTERFACE, 'loadScript', scriptPath, PLUGIN_NAME);
		if (typeof scriptId !== 'number' || scriptId < 0) {
			throw new Error(`KWin refused to load the script (id ${scriptId})`);
		}

		// Plasma 6 exposes `/Scripting/Script<id>`, Plasma 5 `/<id>`.
		let started = false;
		for (const scriptObjectPath of [`${SCRIPTING_PATH}/Script${scriptId}`, `/${scriptId}`]) {
			try {
				// eslint-disable-next-line no-await-in-loop
				await callMethod(KWIN, scriptObjectPath, SCRIPT_INTERFACE, 'run');
				started = true;
				break;
			} catch {}
		}

		if (!started) {
			throw new Error('KWin loaded the script but could not run it');
		}

		this.ready = true;
		this.scriptPath = scriptPath;
		this.registerCleanup();
	}

	registerCleanup() {
		if (this.cleanupRegistered) {
			return;
		}

		this.cleanupRegistered = true;
		process.once('exit', () => {
			// Best effort: the bus write may not flush, the next start unloads by name anyway.
			(async () => {
				try {
					await callMethod(KWIN, SCRIPTING_PATH, SCRIPTING_INTERFACE, 'unloadScript', PLUGIN_NAME);
				} catch {}
			})();

			try {
				fs.unlinkSync(this.scriptPath);
			} catch {}
		});
	}

	onReport(json) {
		let payload;
		try {
			payload = JSON.parse(json);
		} catch {
			payload = null;
		}

		this.last = payload;
		this.lastReportAt = Date.now();
		for (const resolve of this.waiters.splice(0)) {
			resolve();
		}
	}

	waitForFirstReport() {
		if (this.lastReportAt > 0) {
			return Promise.resolve();
		}

		return new Promise(resolve => {
			const timer = setTimeout(resolve, FIRST_REPORT_TIMEOUT);
			this.waiters.push(() => {
				clearTimeout(timer);
				resolve();
			});
		});
	}
}

const bridge = new KWinBridge();

export async function activeWindow() {
	await bridge.ensure();
	await bridge.waitForFirstReport();

	const window = bridge.last;
	if (!window) {
		return undefined;
	}

	return enrichWithProcess(buildResult({
		backend: name,
		title: window.caption,
		id: undefined, // KWin ids are UUIDs, not numbers
		internalId: window.internalId,
		appId: window.resourceClass || window.desktopFileName || window.resourceName,
		name: window.resourceClass || window.desktopFileName || window.resourceName,
		processId: window.pid,
		bounds: window,
	}));
}
