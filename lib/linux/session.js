import process from 'node:process';

const includesAny = (haystack, needles) => needles.some(needle => haystack.includes(needle));

/**
Describe the Linux graphical session from the environment.

Wayland gives us no portable "active window" API, so every backend choice
starts here: which compositor is running and whether an X11 (XWayland)
display is available as a fallback.
*/
/**
Electron (Chromium) rewrites `XDG_CURRENT_DESKTOP` to "Unity" inside its own
processes so GTK shows app-indicator tray icons, and keeps the real value in
`ORIGINAL_XDG_CURRENT_DESKTOP`. Session-specific variables are consulted as
well, since a few launchers set no XDG variable at all.
*/
export function detectDesktop(env = process.env) {
	const xdg = env.ORIGINAL_XDG_CURRENT_DESKTOP || env.XDG_CURRENT_DESKTOP || '';
	const parts = [xdg, env.DESKTOP_SESSION || ''];
	if (env.GNOME_SHELL_SESSION_MODE || env.GNOME_DESKTOP_SESSION_ID) {
		parts.push('GNOME');
	}

	if (env.KDE_FULL_SESSION || env.KDE_SESSION_VERSION) {
		parts.push('KDE');
	}

	return parts.filter(Boolean).join(':');
}

export function detectSession(env = process.env) {
	const desktopRaw = detectDesktop(env);
	const desktops = desktopRaw.toLowerCase().split(':').filter(Boolean);
	const hasWaylandDisplay = Boolean(env.WAYLAND_DISPLAY);
	const hasX11Display = Boolean(env.DISPLAY);

	let sessionType = (env.XDG_SESSION_TYPE || '').toLowerCase();
	if (sessionType !== 'wayland' && sessionType !== 'x11') {
		// `tty`, empty, or something exotic: fall back to the display sockets.
		sessionType = hasWaylandDisplay ? 'wayland' : (hasX11Display ? 'x11' : 'unknown');
	}

	const isWayland = sessionType === 'wayland' || (hasWaylandDisplay && sessionType !== 'x11');

	return {
		sessionType: isWayland ? 'wayland' : sessionType,
		isWayland,
		hasX11Display,
		desktop: desktopRaw,
		isGnome: includesAny(desktops, ['gnome', 'ubuntu', 'pop']),
		isKde: includesAny(desktops, ['kde', 'plasma']),
		isHyprland: Boolean(env.HYPRLAND_INSTANCE_SIGNATURE) || includesAny(desktops, ['hyprland']),
		isSway: Boolean(env.SWAYSOCK) || includesAny(desktops, ['sway']),
	};
}
