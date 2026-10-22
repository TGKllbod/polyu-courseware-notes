import Foundation
import Vision
import AppKit

// Usage: swift vision_ocr.swift <imagePath...>
// Prints recognized text lines per image, separated by markers.

let args = CommandLine.arguments
guard args.count > 1 else {
    print("no args")
    exit(1)
}

for path in args.dropFirst() {
    guard let img = NSImage(contentsOfFile: path),
          let cg = img.cgImage(forProposedRect: nil, context: nil, hints: nil) else {
        print("=== \(path) === [LOAD FAIL]")
        continue
    }
    let request = VNRecognizeTextRequest()
    request.recognitionLevel = .accurate
    request.usesLanguageCorrection = true
    request.recognitionLanguages = ["en-US", "zh-Hans", "zh-Hant"]
    let handler = VNImageRequestHandler(cgImage: cg, options: [:])
    do {
        try handler.perform([request])
        let observations = request.results ?? []
        print("=== \(path) ===")
        if observations.isEmpty {
            print("[NO TEXT]")
        } else {
            for obs in observations {
                if let cand = obs.topCandidates(1).first {
                    print(cand.string)
                }
            }
        }
    } catch {
        print("=== \(path) === [ERROR \(error)]")
    }
}
