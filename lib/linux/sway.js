import {promisify} from 'node:util';
import childProcess from 'node:child_process';
import {
	buildResult,
	enrichWithProcess,
	enrichWithProcessSync,
} from './process-info.js';

/**
Sway backend: walk `swaymsg -t get_tree` for the focused container.
*/

const execFile = promisify(childProcess.execFile);

export const name = 'sway';

const isWindowNode = node => (node.type === 'con' || node.type === 'floating_con') && (node.app_id || node.window_properties || node.pid);

export function collectWindows(tree) {
	const windows = [];
	const stack = [tree];
	while (stack.length > 0) {
		const node = stack.pop();
		if (!node || typeof node !== 'object') {
			continue;
		}

		if (isWindowNode(node)) {
			windows.push(node);
		}

		stack.push(...(node.nodes ?? []), ...(node.floating_nodes ?? []));
	}

	return windows;
}

export const findFocused = tree => collectWindows(tree).find(node => node.focused);

const fromNode = node => {
	if (!node) {
		return undefined;
	}

	// XWayland clients carry X11 WM_CLASS instead of app_id.
	const appId = node.app_id || node.window_properties?.class || node.window_properties?.instance;
	return buildResult({
		backend: name,
		title: node.name,
		id: node.id,
		appId,
		name: appId,
		processId: node.pid,
		bounds: node.rect,
	});
};

const parse = stdout => {
	try {
		return JSON.parse(stdout);
	} catch {
		return undefined;
	}
};

export async function activeWindow() {
	const {stdout} = await execFile('swaymsg', ['-t', 'get_tree', '-r']);
	const result = fromNode(findFocused(parse(stdout)));
	return result && enrichWithProcess(result);
}

export function activeWindowSync() {
	const stdout = childProcess.execFileSync('swaymsg', ['-t', 'get_tree', '-r'], {encoding: 'utf8'});
	const result = fromNode(findFocused(parse(stdout)));
	return result && enrichWithProcessSync(result);
}

export async function openWindows() {
	const {stdout} = await execFile('swaymsg', ['-t', 'get_tree', '-r']);
	return Promise.all(collectWindows(parse(stdout)).map(node => enrichWithProcess(fromNode(node))));
}

export function openWindowsSync() {
	const stdout = childProcess.execFileSync('swaymsg', ['-t', 'get_tree', '-r'], {encoding: 'utf8'});
	return collectWindows(parse(stdout)).map(node => enrichWithProcessSync(fromNode(node)));
}
