import AppKit


@discardableResult
func runAppleScript(source: String) -> String? {
	NSAppleScript(source: source)?.executeAndReturnError(nil).stringValue
}


func toJson<T>(_ data: T) throws -> String {
	let json = try JSONSerialization.data(withJSONObject: data)
	return String(data: json, encoding: .utf8)!
}


// Show the system prompt if there's no permission.
// Ask TCC rather than open a display stream: creating one can hang in WindowServer for about 30 s and
// then return nil with the permission granted, which read as a missing permission.
func hasScreenRecordingPermission() -> Bool {
	CGPreflightScreenCaptureAccess() || CGRequestScreenCaptureAccess()
}
