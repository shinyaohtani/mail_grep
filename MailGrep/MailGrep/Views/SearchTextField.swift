import AppKit
import SwiftUI

/// コンテクストメニューに「MailGrepで検索...」を含むカスタムテキストフィールド
struct SearchTextField: NSViewRepresentable {
    @Binding var text: String
    var placeholder: String
    var onSubmit: () -> Void

    func makeNSView(context: Context) -> NSTextField {
        let textField = NSTextField()
        textField.stringValue = text  // 初期値を設定
        textField.placeholderString = placeholder
        textField.delegate = context.coordinator
        textField.bezelStyle = .roundedBezel
        textField.focusRingType = .exterior
        textField.setAccessibilityIdentifier("searchPatternTextField")
        textField.setAccessibilityLabel("検索パターン")
        return textField
    }

    func updateNSView(_ nsView: NSTextField, context: Context) {
        // Coordinatorのparent参照を更新（SwiftUIがビューを再作成した場合に必要）
        context.coordinator.parent = self

        // delegateが失われていないか確認し、再設定
        if nsView.delegate !== context.coordinator {
            nsView.delegate = context.coordinator
        }
        if nsView.stringValue != text {
            nsView.stringValue = text
        }
    }

    func makeCoordinator() -> Coordinator {
        Coordinator(self)
    }

    class Coordinator: NSObject, NSTextFieldDelegate {
        var parent: SearchTextField

        init(_ parent: SearchTextField) {
            self.parent = parent
        }

        func controlTextDidChange(_ obj: Notification) {
            if let textField = obj.object as? NSTextField {
                parent.text = textField.stringValue
            }
        }

        func control(_: NSControl, textView _: NSTextView, doCommandBy commandSelector: Selector) -> Bool {
            if commandSelector == #selector(NSResponder.insertNewline(_:)) {
                parent.onSubmit()
                return true
            }
            return false
        }

        /// フィールドエディタが開始されたときにコンテクストメニューをカスタマイズ
        func controlTextDidBeginEditing(_ obj: Notification) {
            guard let textField = obj.object as? NSTextField,
                  let editor = textField.currentEditor() as? NSTextView
            else { return }

            // フィールドエディタのメニューをカスタマイズ
            setupFieldEditorMenu(editor, textField: textField)
        }

        private func setupFieldEditorMenu(_ editor: NSTextView, textField: NSTextField) {
            // 既存のメニューを取得またはデフォルトメニューを作成
            let menu = editor.menu ?? NSMenu()

            // 既に追加済みかチェック
            if menu.item(withTitle: "MailGrepで検索...") != nil {
                return
            }

            // セパレーターを追加
            menu.addItem(NSMenuItem.separator())

            // 「MailGrepで検索...」メニュー項目を追加
            let mailGrepItem = NSMenuItem(
                title: "MailGrepで検索...",
                action: #selector(searchWithMailGrep(_:)),
                keyEquivalent: ""
            )
            mailGrepItem.target = self
            mailGrepItem.representedObject = textField
            menu.addItem(mailGrepItem)

            editor.menu = menu
        }

        @objc func searchWithMailGrep(_ sender: NSMenuItem) {
            guard let textField = sender.representedObject as? NSTextField,
                  let editor = textField.currentEditor() as? NSTextView
            else { return }

            // 選択されているテキストを取得
            let selectedText: String
            if let range = editor.selectedRanges.first as? NSRange, range.length > 0 {
                selectedText = (editor.string as NSString).substring(with: range)
            } else {
                // 選択がない場合は全テキストを使用
                selectedText = textField.stringValue
            }

            guard !selectedText.isEmpty else { return }

            // 検索を実行するためにNotificationを送信
            NotificationCenter.default.post(
                name: .contextMenuSearchRequested,
                object: selectedText
            )
        }
    }
}

extension Notification.Name {
    static let contextMenuSearchRequested = Notification.Name("contextMenuSearchRequested")
}
