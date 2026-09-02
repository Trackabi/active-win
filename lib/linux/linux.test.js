/* eslint-disable camelcase -- sway's JSON is snake_case by protocol */
import test from 'ava';
import {candidateBackends} from '../linux.js';
import {detectSession} from './session.js';
import {normalizeAppName, buildResult} from './process-info.js';
import {findFocused, collectWindows} from './sway.js';
import {parseLinux} from './x11.js';

test('detectSession: GNOME Wayland with XWayland display', t => {
	const session = detectSession({
		XDG_SESSION_TYPE: 'wayland',
		WAYLAND_DISPLAY: 'wayland-0',
		DISPLAY: ':0',
		XDG_CURRENT_DESKTOP: 'ubuntu:GNOME',
	});
	t.true(session.isWayland);
	t.true(session.isGnome);
	t.false(session.isKde);
	t.true(session.hasX11Display);
	t.is(session.sessionType, 'wayland');
});

test('detectSession: falls back to display sockets when XDG_SESSION_TYPE is tty', t => {
	t.true(detectSession({XDG_SESSION_TYPE: 'tty', WAYLAND_DISPLAY: 'wayland-1', XDG_CURRENT_DESKTOP: 'KDE'}).isWayland);
	t.is(detectSession({XDG_SESSION_TYPE: 'tty', DISPLAY: ':1'}).sessionType, 'x11');
	t.is(detectSession({}).sessionType, 'unknown');
});

test('detectSession: Hyprland and Sway from their sockets', t => {
	t.true(detectSession({WAYLAND_DISPLAY: 'wayland-1', HYPRLAND_INSTANCE_SIGNATURE: 'abc'}).isHyprland);
	t.true(detectSession({WAYLAND_DISPLAY: 'wayland-1', SWAYSOCK: '/run/user/1000/sway.sock'}).isSway);
});

test('normalizeAppName: reverse-DNS ids collapse to the last segment', t => {
	t.is(normalizeAppName('org.mozilla.firefox'), 'firefox');
	t.is(normalizeAppName('org.gnome.Nautilus'), 'Nautilus');
	t.is(normalizeAppName('com.google.Chrome.desktop'), 'Chrome');
	t.is(normalizeAppName('firefox'), 'firefox');
	t.is(normalizeAppName('Code'), 'Code');
	t.is(normalizeAppName('google-chrome'), 'google-chrome');
	t.is(normalizeAppName('Some App 1.2'), 'Some App 1.2');
	t.is(normalizeAppName(undefined), '');
});

test('buildResult: coerces ids and keeps the raw app id', t => {
	const result = buildResult({
		backend: 'hyprland',
		title: 'Inbox',
		id: '0x1a',
		appId: 'org.mozilla.firefox',
		name: 'org.mozilla.firefox',
		processId: '4242',
		bounds: {
			x: '10', y: 20, width: undefined, height: 30,
		},
	});
	t.is(result.platform, 'linux');
	t.is(result.backend, 'hyprland');
	t.is(result.owner.name, 'firefox');
	t.is(result.owner.appId, 'org.mozilla.firefox');
	t.is(result.owner.processId, 4242);
	t.deepEqual(result.bounds, {
		x: 10, y: 20, width: 0, height: 30,
	});
});

test('sway: finds the focused container in a nested tree', t => {
	const tree = {
		type: 'root',
		nodes: [{
			type: 'output',
			nodes: [{
				type: 'workspace',
				nodes: [
					{
						type: 'con', id: 1, app_id: 'foot', name: 'shell', pid: 10, focused: false, rect: {
							x: 0, y: 0, width: 100, height: 100,
						},
					},
					{
						type: 'con', id: 2, app_id: null, window_properties: {class: 'Steam', instance: 'steam'}, name: 'Steam', pid: 11, focused: true, rect: {
							x: 100, y: 0, width: 100, height: 100,
						},
					},
				],
				floating_nodes: [
					{
						type: 'floating_con', id: 3, app_id: 'pavucontrol', name: 'Volume', pid: 12, focused: false,
					},
				],
			}],
		}],
	};
	t.is(collectWindows(tree).length, 3);
	t.is(findFocused(tree).id, 2);
});

test('x11: parses xprop output and survives a missing WM_CLASS', t => {
	const stdout = [
		'_NET_WM_NAME(UTF8_STRING) = "Editor – file.js"',
		'WM_CLASS(STRING) = "code", "Code"',
		'_NET_WM_PID(CARDINAL) = 1234',
	].join('\n');
	const boundsStdout = 'Absolute upper-left X:  5\nAbsolute upper-left Y:  6\nWidth: 700\nHeight: 500';
	const result = parseLinux({stdout, boundsStdout, activeWindowId: 99});
	t.is(result.title, 'Editor – file.js');
	t.is(result.owner.name, 'Code');
	t.is(result.owner.processId, 1234);
	t.is(result.id, 99);
	t.deepEqual(result.bounds, {
		x: 5, y: 6, width: 700, height: 500,
	});

	const noClass = parseLinux({stdout: 'WM_NAME(STRING) = "x"', boundsStdout: '', activeWindowId: 7});
	t.is(noClass.owner.name, '');
	t.is(noClass.title, 'x');
});

test('candidateBackends: X11 session only probes xprop', t => {
	t.deepEqual(candidateBackends(detectSession({XDG_SESSION_TYPE: 'x11', DISPLAY: ':0'})).map(b => b.name), ['x11']);
});

test('candidateBackends: Wayland sessions end with the X11 fallback', t => {
	const gnome = candidateBackends(detectSession({XDG_SESSION_TYPE: 'wayland', DISPLAY: ':0', XDG_CURRENT_DESKTOP: 'GNOME'})).map(b => b.name);
	t.deepEqual(gnome, ['gnome-focused-window-dbus', 'gnome-window-calls-extended', 'gnome-window-calls', 'x11']);

	const kde = candidateBackends(detectSession({XDG_SESSION_TYPE: 'wayland', DISPLAY: ':0', XDG_CURRENT_DESKTOP: 'KDE'})).map(b => b.name);
	t.deepEqual(kde, ['kwin', 'x11']);

	const unknownNoX = new Set(candidateBackends(detectSession({XDG_SESSION_TYPE: 'wayland', XDG_CURRENT_DESKTOP: 'river'})).map(b => b.name));
	t.true(unknownNoX.has('hyprland') && unknownNoX.has('kwin'));
	t.false(unknownNoX.has('x11'));
});
