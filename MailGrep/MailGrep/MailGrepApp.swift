import SwiftUI

@main
struct MailGrepApp: App {
    @NSApplicationDelegateAdaptor(AppDelegate.self) var appDelegate

    var body: some Scene {
        WindowGroup(id: "search") {
            ContentView()
        }
        .commands {
            // 新規ウィンドウメニュー - OpenWindowActionを使用
            CommandGroup(replacing: .newItem) {
                OpenNewWindowButton()
            }
        }
    }
}

/// 新規ウィンドウを開くボタン（@EnvironmentをView内で使用するため）
struct OpenNewWindowButton: View {
    @Environment(\.openWindow) private var openWindow

    var body: some View {
        Button("新規検索ウィンドウ") {
            openWindow(id: "search")
        }
        .keyboardShortcut("n", modifiers: .command)
    }
}

extension Notification.Name {
    static let serviceSearchRequested = Notification.Name("serviceSearchRequested")
    static let searchHistoryUpdated = Notification.Name("searchHistoryUpdated")
}
