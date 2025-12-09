import AppKit
import SwiftUI

struct ContentView: View {
    @EnvironmentObject var viewModel: SearchViewModel

    var body: some View {
        VStack(spacing: 0) {
            searchBar
            Divider()
            resultStats
            ResultsTableView(results: viewModel.results) { hitLine in
                viewModel.openMailInApp(hitLine)
            }
            Divider()
            bottomBar
        }
        .frame(minWidth: 800, minHeight: 500)
    }

    private var searchBar: some View {
        HStack(spacing: 12) {
            Text("検索パターン:")
                .foregroundColor(.secondary)

            SearchFieldWithHistory(
                text: $viewModel.pattern,
                placeholder: "正規表現を入力...",
                history: viewModel.searchHistory,
                isDisabled: viewModel.isSearching,
                onSubmit: { viewModel.search() },
                onSelectHistory: { keyword in
                    viewModel.selectFromHistory(keyword)
                },
                onClearHistory: {
                    viewModel.clearHistory()
                }
            )
            .frame(minWidth: 200)

            if viewModel.isSearching {
                Button(action: {
                    viewModel.cancelSearch()
                }) {
                    HStack(spacing: 4) {
                        Image(systemName: "stop.fill")
                        Text("停止")
                    }
                }
                .keyboardShortcut(.escape, modifiers: [])
            } else {
                Button(action: {
                    viewModel.search()
                }) {
                    HStack(spacing: 4) {
                        Image(systemName: "magnifyingglass")
                        Text("検索")
                    }
                }
                .keyboardShortcut(.return, modifiers: .command)
                .disabled(viewModel.pattern.isEmpty)
            }
        }
        .padding()
    }

    private var searchOptions: some View {
        HStack(spacing: 20) {
            Toggle("大文字小文字を無視", isOn: $viewModel.ignoreCase)
            Toggle("送信済みのみ", isOn: $viewModel.onlySent)
            Spacer()
        }
        .padding(.horizontal)
        .padding(.bottom, 8)
    }

    private var resultStats: some View {
        HStack {
            if viewModel.isSearching {
                ProgressView()
                    .scaleEffect(0.7)
                Text(viewModel.statusMessage)
                    .foregroundColor(.secondary)
            } else if !viewModel.results.isEmpty {
                Text("結果: \(viewModel.results.count)件 (\(viewModel.mailCount)通のメール)")
                    .foregroundColor(.secondary)
            } else if viewModel.searchCompleted {
                Text("検索結果なし")
                    .foregroundColor(.secondary)
            }
            Spacer()

            Toggle("大文字小文字を無視", isOn: $viewModel.ignoreCase)
                .toggleStyle(.checkbox)
            Toggle("送信済みのみ", isOn: $viewModel.onlySent)
                .toggleStyle(.checkbox)
        }
        .padding(.horizontal)
        .padding(.vertical, 8)
    }

    private var bottomBar: some View {
        HStack {
            Button("CSV保存") {
                viewModel.exportCSV()
            }
            .disabled(viewModel.results.isEmpty)

            Spacer()

            if viewModel.isSearching {
                ProgressView(value: viewModel.progress)
                    .frame(width: 200)
                Text("\(Int(viewModel.progress * 100))%")
                    .foregroundColor(.secondary)
                    .frame(width: 40)
            }
        }
        .padding()
    }
}

// MARK: - Search Field with History

/// 検索履歴ドロップダウン付きの検索フィールド
struct SearchFieldWithHistory: View {
    @Binding var text: String
    var placeholder: String
    var history: [String]
    var isDisabled: Bool
    var onSubmit: () -> Void
    var onSelectHistory: (String) -> Void
    var onClearHistory: () -> Void

    @State private var isShowingHistory = false

    var body: some View {
        HStack(spacing: 0) {
            // 虫眼鏡 + v アイコンボタン
            Button(action: {
                isShowingHistory.toggle()
            }) {
                HStack(spacing: 2) {
                    Image(systemName: "magnifyingglass")
                        .foregroundColor(.secondary)
                    Image(systemName: "chevron.down")
                        .font(.system(size: 8))
                        .foregroundColor(.secondary)
                }
                .padding(.horizontal, 6)
                .padding(.vertical, 4)
                .contentShape(Rectangle())
            }
            .buttonStyle(.plain)
            .disabled(isDisabled)
            .popover(isPresented: $isShowingHistory, arrowEdge: .bottom) {
                historyMenu
            }

            // テキストフィールド
            TextField(placeholder, text: $text)
                .textFieldStyle(.plain)
                .disabled(isDisabled)
                .onSubmit {
                    onSubmit()
                }
        }
        .padding(.horizontal, 4)
        .padding(.vertical, 2)
        .background(isDisabled ? Color(NSColor.controlBackgroundColor) : Color(NSColor.textBackgroundColor))
        .cornerRadius(6)
        .overlay(
            RoundedRectangle(cornerRadius: 6)
                .stroke(Color(NSColor.separatorColor), lineWidth: 1)
        )
    }

    @ViewBuilder
    private var historyMenu: some View {
        VStack(alignment: .leading, spacing: 0) {
            if history.isEmpty {
                Text("最近の検索はありません")
                    .foregroundColor(.secondary)
                    .padding(.horizontal, 12)
                    .padding(.vertical, 8)
            } else {
                Text("最近の検索")
                    .font(.caption)
                    .foregroundColor(.secondary)
                    .padding(.horizontal, 12)
                    .padding(.top, 8)
                    .padding(.bottom, 4)

                Divider()

                ScrollView {
                    VStack(alignment: .leading, spacing: 0) {
                        ForEach(history, id: \.self) { keyword in
                            Button(action: {
                                isShowingHistory = false
                                onSelectHistory(keyword)
                            }) {
                                HStack {
                                    Image(systemName: "clock.arrow.circlepath")
                                        .foregroundColor(.secondary)
                                        .font(.system(size: 12))
                                    Text(keyword)
                                        .lineLimit(1)
                                        .truncationMode(.tail)
                                    Spacer()
                                }
                                .padding(.horizontal, 12)
                                .padding(.vertical, 6)
                                .contentShape(Rectangle())
                            }
                            .buttonStyle(.plain)
                        }
                    }
                }
                .frame(maxHeight: 200)

                Divider()

                Button(action: {
                    isShowingHistory = false
                    onClearHistory()
                }) {
                    HStack {
                        Image(systemName: "trash")
                            .foregroundColor(.red)
                            .font(.system(size: 12))
                        Text("履歴をクリア")
                            .foregroundColor(.red)
                    }
                    .padding(.horizontal, 12)
                    .padding(.vertical, 8)
                    .contentShape(Rectangle())
                }
                .buttonStyle(.plain)
            }
        }
        .frame(minWidth: 200)
    }
}

#Preview {
    ContentView()
        .environmentObject(SearchViewModel())
}
