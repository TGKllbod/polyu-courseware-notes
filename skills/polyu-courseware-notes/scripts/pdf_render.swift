import Foundation
import PDFKit
import AppKit

// Usage: swift pdf_render.swift <pdfPath> <outDir> [scale] [pageList]
//   pageList: e.g. "7,20,39,60" or "all" (default all)
// Renders each page to <outDir>/pNNN.png at the given scale (default 3.0 ≈ 216 dpi).
// Used to feed image-only slides (charts, screenshots, cartoons) into vision_ocr.swift.

let args = CommandLine.arguments
guard args.count > 2 else {
    print("usage: swift pdf_render.swift <pdfPath> <outDir> [scale] [pageList|all]")
    exit(1)
}
let pdfPath = args[1]
let outDir = args[2]
let scale = args.count > 3 ? (Double(args[3]) ?? 3.0) : 3.0
let pageList = args.count > 4 ? args[4] : "all"

guard let doc = PDFDocument(url: URL(fileURLWithPath: pdfPath)) else {
    print("[LOAD FAIL] \(pdfPath)")
    exit(1)
}
try? FileManager.default.createDirectory(atPath: outDir, withIntermediateDirectories: true)

var wanted: [Int] = []
if pageList.lowercased() == "all" {
    wanted = Array(0..<doc.pageCount)
} else {
    wanted = pageList.split(separator: ",").compactMap { Int($0.trimmingCharacters(in: .whitespaces)) }.map { $0 - 1 }
}

for idx in wanted {
    guard idx >= 0, idx < doc.pageCount, let page = doc.page(at: idx) else {
        print("p\(idx + 1): [OUT OF RANGE]")
        continue
    }
    let bounds = page.bounds(for: .mediaBox)
    let w = Int(bounds.width * scale)
    let h = Int(bounds.height * scale)
    guard let rep = NSBitmapImageRep(bitmapDataPlanes: nil, pixelsWide: w, pixelsHigh: h,
                                    bitsPerSample: 8, samplesPerPixel: 4, hasAlpha: true,
                                    isPlanar: false, colorSpaceName: .deviceRGB,
                                    bytesPerRow: 0, bitsPerPixel: 0) else {
        print("p\(idx + 1): [BITMAP FAIL]")
        continue
    }
    guard let ctx = NSGraphicsContext(bitmapImageRep: rep) else {
        print("p\(idx + 1): [CTX FAIL]")
        continue
    }
    NSGraphicsContext.saveGraphicsState()
    NSGraphicsContext.current = ctx
    NSColor.white.setFill()
    NSRect(x: 0, y: 0, width: w, height: h).fill()
    ctx.cgContext.scaleBy(x: scale, y: scale)
    page.draw(with: .mediaBox, to: ctx.cgContext)
    NSGraphicsContext.restoreGraphicsState()

    guard let png = rep.representation(using: .png, properties: [:]) else {
        print("p\(idx + 1): [PNG FAIL]")
        continue
    }
    let outPath = "\(outDir)/p\(String(format: "%03d", idx + 1)).png"
    do {
        try png.write(to: URL(fileURLWithPath: outPath))
        print("p\(idx + 1): \(outPath)  \(w)x\(h)")
    } catch {
        print("p\(idx + 1): [WRITE FAIL] \(error)")
    }
}
