import SwiftUI

struct ResultsTableView: View {
    let results: [HitLine]
    let onRowClick: (HitLine) -> Void

    @State private var selection: HitLine.ID?
    @State private var sortOrder = [KeyPathComparator(\HitLine.mailID)]

    var body: some View {
        Table(results, selection: $selection, sortOrder: $sortOrder) {
            TableColumn("#", value: \.mailID) { hitLine in
                Text("\(hitLine.mailID)")
                    .frame(maxWidth: .infinity, alignment: .trailing)
            }
            .width(min: 30, ideal: 40, max: 60)

            TableColumn("日付", value: \.dateStr) { hitLine in
                Text(hitLine.dateStr)
            }
            .width(min: 100, ideal: 140, max: 180)

            TableColumn("件名", value: \.subject) { hitLine in
                Text(hitLine.subject)
                    .lineLimit(1)
            }
            .width(min: 150, ideal: 250)

            TableColumn("From", value: \.fromAddr) { hitLine in
                Text(hitLine.fromAddr)
                    .lineLimit(1)
            }
            .width(min: 100, ideal: 180)

            TableColumn("マッチ行", value: \.matchedLine) { hitLine in
                Text(hitLine.matchedLine)
                    .lineLimit(1)
                    .foregroundColor(.secondary)
            }
            .width(min: 150, ideal: 300)
        }
        .onChange(of: selection) { newValue in
            if let id = newValue, let hitLine = results.first(where: { $0.id == id }) {
                onRowClick(hitLine)
                selection = nil
            }
        }
        .contextMenu(forSelectionType: HitLine.ID.self) { ids in
            if let id = ids.first, let hitLine = results.first(where: { $0.id == id }) {
                Button("Mail.appで開く") {
                    onRowClick(hitLine)
                }
                Divider()
                Button("件名をコピー") {
                    NSPasteboard.general.clearContents()
                    NSPasteboard.general.setString(hitLine.subject, forType: .string)
                }
                Button("マッチ行をコピー") {
                    NSPasteboard.general.clearContents()
                    NSPasteboard.general.setString(hitLine.matchedLine, forType: .string)
                }
            }
        } primaryAction: { ids in
            if let id = ids.first, let hitLine = results.first(where: { $0.id == id }) {
                onRowClick(hitLine)
            }
        }
    }
}

#Preview {
    ResultsTableView(results: []) { _ in }
}
