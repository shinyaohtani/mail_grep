import SwiftUI

@main
struct MailGrepApp: App {
    @NSApplicationDelegateAdaptor(AppDelegate.self) var appDelegate
    @StateObject private var searchViewModel = SearchViewModel()

    var body: some Scene {
        WindowGroup {
            ContentView()
                .environmentObject(searchViewModel)
                .onReceive(NotificationCenter.default.publisher(for: .serviceSearchRequested)) { notification in
                    if let text = notification.object as? String {
                        searchViewModel.pattern = text
                        searchViewModel.search()
                    }
                }
        }
        .commands {
            CommandGroup(replacing: .newItem) {}
        }
    }
}

extension Notification.Name {
    static let serviceSearchRequested = Notification.Name("serviceSearchRequested")
}
