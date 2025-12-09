import AppKit

class AppDelegate: NSObject, NSApplicationDelegate {
    let serviceProvider = ServiceProvider()

    func applicationDidFinishLaunching(_: Notification) {
        NSApp.servicesProvider = serviceProvider
        NSUpdateDynamicServices()

        // 起動時にフルディスクアクセス権限をチェック
        PermissionChecker.shared.showPermissionAlertIfNeeded()
    }

    func applicationShouldTerminateAfterLastWindowClosed(_: NSApplication) -> Bool {
        true
    }

    func applicationWillTerminate(_: Notification) {
        // 通常終了時にViewModelに通知
        NotificationCenter.default.post(name: .appWillTerminateNormally, object: nil)
    }
}

extension Notification.Name {
    static let appWillTerminateNormally = Notification.Name("appWillTerminateNormally")
}

class ServiceProvider: NSObject {
    @objc func searchWithText(
        _ pasteboard: NSPasteboard,
        userData _: String?,
        error _: AutoreleasingUnsafeMutablePointer<NSString?>
    ) {
        guard let text = pasteboard.string(forType: .string) else { return }
        NSApp.activate(ignoringOtherApps: true)
        NotificationCenter.default.post(
            name: .serviceSearchRequested,
            object: text
        )
    }
}

// MARK: - Permission Checker

/// フルディスクアクセス権限をチェックするサービス
class PermissionChecker {
    static let shared = PermissionChecker()

    private init() {}

    /// ~/Library/Mail へのアクセス権限があるかチェック
    func hasMailAccess() -> Bool {
        let homeDir = FileManager.default.homeDirectoryForCurrentUser
        let mailDir = homeDir.appendingPathComponent("Library/Mail")

        // ディレクトリの存在確認
        var isDirectory: ObjCBool = false
        guard FileManager.default.fileExists(atPath: mailDir.path, isDirectory: &isDirectory),
              isDirectory.boolValue
        else {
            // Mailディレクトリが存在しない（Mail.appを使っていない可能性）
            return true
        }

        // ディレクトリの内容を読み取れるかチェック
        do {
            _ = try FileManager.default.contentsOfDirectory(at: mailDir, includingPropertiesForKeys: nil)
            // 内容が読み取れれば権限あり
            return true
        } catch {
            // 権限エラーの場合はfalse
            return false
        }
    }

    /// 現在のアプリのパスを取得
    func currentAppPath() -> String {
        Bundle.main.bundlePath
    }

    /// 権限がない場合にアラートを表示
    func showPermissionAlertIfNeeded() {
        guard !hasMailAccess() else { return }

        DispatchQueue.main.async {
            let alert = NSAlert()
            alert.messageText = "フルディスクアクセス権限が必要です"
            alert.informativeText = """
            MailGrepがメールデータにアクセスするには「フルディスクアクセス」権限が必要です。

            以下の手順で権限を付与してください：
            1. 「システム設定を開く」をクリック
            2. 「フルディスクアクセス」を選択
            3. 「+」をクリックしてこのアプリを追加：
               \(self.currentAppPath())
            4. アプリを再起動

            ※ 複数のMailGrep.appが存在する場合は、現在使用しているアプリを追加してください。
            """
            alert.alertStyle = .warning
            alert.addButton(withTitle: "システム設定を開く")
            alert.addButton(withTitle: "後で")

            let response = alert.runModal()
            if response == .alertFirstButtonReturn {
                self.openPrivacySettings()
            }
        }
    }

    /// プライバシー設定を開く
    func openPrivacySettings() {
        // macOS 13+ ではこのURLでフルディスクアクセス設定を直接開ける
        if let url = URL(string: "x-apple.systempreferences:com.apple.preference.security?Privacy_AllFiles") {
            NSWorkspace.shared.open(url)
        }
    }
}
