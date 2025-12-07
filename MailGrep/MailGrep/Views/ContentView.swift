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

            TextField("正規表現を入力...", text: $viewModel.pattern)
                .textFieldStyle(.roundedBorder)
                .onSubmit {
                    viewModel.search()
                }

            Button(action: { viewModel.search() }) {
                HStack(spacing: 4) {
                    Image(systemName: "magnifyingglass")
                    Text("検索")
                }
            }
            .keyboardShortcut(.return, modifiers: .command)
            .disabled(viewModel.isSearching || viewModel.pattern.isEmpty)
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

#Preview {
    ContentView()
        .environmentObject(SearchViewModel())
}
