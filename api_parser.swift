//
//  api_parser.swift
//  PrivCite v2.0 - Core Engine JSON API Bridge
//
//  This command-line tool bridges the pure Swift PrivCiteEngine to the web-based
//  Word JS Sidebar. It reads plain text from stdin, runs a full analysis, and
//  outputs the resulting DisplayIssues as a serialized JSON array.
//

import Foundation

#if canImport(UniformTypeIdentifiers)
import UniformTypeIdentifiers
#endif

// Define a Codable representation of DisplayIssue for JSON serialization
struct JSONDisplayIssue: Codable {
    let id: String
    let category: String
    let severity: String
    let message: String
    let targetText: String?
    let suggestion: String
    let location: String
}

@main
struct APIParser {
    static func main() {
        // 1. Read all text from standard input (stdin)
        var inputData = Data()
        let stdin = FileHandle.standardInput
        
        while true {
            let data = stdin.availableData
            if data.isEmpty {
                break
            }
            inputData.append(data)
        }
        
        guard let text = String(data: inputData, encoding: .utf8), !text.isEmpty else {
            print("[]") // Return empty JSON array if no text provided
            return
        }
        
        
        // Read spacing preference from arguments (passed by Python server)
        var spacingPref: SpacingPreference = .adaptive
        var disableSpacing = false
        let args = CommandLine.arguments
        if let prefIndex = args.firstIndex(of: "--spacing-pref"), prefIndex + 1 < args.count {
            let val = args[prefIndex + 1]
            if val == "disabled" {
                disableSpacing = true
            } else if let parsed = SpacingPreference(rawValue: val) {
                spacingPref = parsed
            }
        }
        PrivCiteEngine.globalSpacingPreference = spacingPref

        // 2. Write the text to a temporary text file
        let tempDir = FileManager.default.temporaryDirectory
        let tempFileURL = tempDir.appendingPathComponent("privcite_api_draft_\(UUID().uuidString).txt")
        
        do {
            try text.write(to: tempFileURL, atomically: true, encoding: .utf8)
        } catch {
            // Output JSON error
            print("[]")
            return
        }
        
        defer {
            try? FileManager.default.removeItem(at: tempFileURL)
        }
        
        // 3. Execute the full production PrivCiteEngine analysis on the text file
        // Enable citations, spacing, grammar, and passive voice checks
        var checks: Set<String> = ["citations", "spacing", "grammar", "passive"]
        if disableSpacing {
            checks.remove("spacing")
        }
        let result = PrivCiteEngine.analyze(
            documentURL: tempFileURL,
            requestedChecks: checks
        )
        
        // 4. Map the engine's DisplayIssues to JSON-friendly structures
        let jsonIssues = result.issues.map { issue -> JSONDisplayIssue in
            return JSONDisplayIssue(
                id: issue.id.uuidString,
                category: issue.category,
                // Match the lowercase severities used by our CSS theme (error, warning, info)
                severity: issue.severity.lowercased(),
                message: issue.message,
                // Map the citationText parameter to targetText for sidebar JS compatibility
                targetText: issue.citationText ?? "",
                suggestion: issue.suggestion,
                location: issue.location
            )
        }
        
        // 5. Serialize to JSON and print to stdout
        let encoder = JSONEncoder()
        encoder.outputFormatting = .prettyPrinted
        
        do {
            let jsonData = try encoder.encode(jsonIssues)
            if let jsonString = String(data: jsonData, encoding: .utf8) {
                print(jsonString)
            } else {
                print("[]")
            }
        } catch {
            print("[]")
        }
    }
}
