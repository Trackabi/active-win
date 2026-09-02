import process from 'node:process';

const includesAny = (haystack, needles) => needles.some(needle => haystack.includes(needle));

/**
Describe the Linux graphical session from the environment.

Wayland gives us no portable "active window" API, so every backend choice
starts here: which compositor is running and whether an X11 (XWayland)
display is available as a fallback.
*/
export function detectSession(env = process.env) {
	const desktopRaw = env.XDG_CURRENT_DESKTOP || env.DESKTOP_SESSION || '';
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
