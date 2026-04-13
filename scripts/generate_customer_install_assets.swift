import AppKit
import CoreImage
import Foundation

struct RGBAImage {
  let width: Int
  let height: Int
  let bytesPerRow: Int
  let pixels: [UInt8]
}

func loadCGImage(at path: String) throws -> CGImage {
  let url = URL(fileURLWithPath: path)
  guard
    let imageSource = CGImageSourceCreateWithURL(url as CFURL, nil),
    let image = CGImageSourceCreateImageAtIndex(imageSource, 0, nil)
  else {
    throw NSError(domain: "AssetGen", code: 1, userInfo: [NSLocalizedDescriptionKey: "Unable to load image at \(path)"])
  }
  return image
}

func rgbaImage(from image: CGImage) throws -> RGBAImage {
  let width = image.width
  let height = image.height
  let bytesPerPixel = 4
  let bytesPerRow = width * bytesPerPixel
  var pixels = [UInt8](repeating: 0, count: height * bytesPerRow)

  let colorSpace = CGColorSpaceCreateDeviceRGB()
  guard
    let context = CGContext(
      data: &pixels,
      width: width,
      height: height,
      bitsPerComponent: 8,
      bytesPerRow: bytesPerRow,
      space: colorSpace,
      bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue
    )
  else {
    throw NSError(domain: "AssetGen", code: 2, userInfo: [NSLocalizedDescriptionKey: "Unable to create RGBA context"])
  }

  context.draw(image, in: CGRect(x: 0, y: 0, width: width, height: height))
  return RGBAImage(width: width, height: height, bytesPerRow: bytesPerRow, pixels: pixels)
}

func nonWhiteBounds(in image: RGBAImage, threshold: UInt8 = 248) -> CGRect {
  var minX = image.width
  var minY = image.height
  var maxX = 0
  var maxY = 0
  var found = false

  for y in 0..<image.height {
    for x in 0..<image.width {
      let offset = y * image.bytesPerRow + x * 4
      let r = image.pixels[offset]
      let g = image.pixels[offset + 1]
      let b = image.pixels[offset + 2]
      let a = image.pixels[offset + 3]
      if a < 8 { continue }
      if r >= threshold && g >= threshold && b >= threshold { continue }
      minX = min(minX, x)
      minY = min(minY, y)
      maxX = max(maxX, x)
      maxY = max(maxY, y)
      found = true
    }
  }

  if !found {
    return CGRect(x: 0, y: 0, width: image.width, height: image.height)
  }

  let padding = max(4, image.width / 64)
  let originX = max(0, minX - padding)
  let originY = max(0, minY - padding)
  let maxBoundX = min(image.width - 1, maxX + padding)
  let maxBoundY = min(image.height - 1, maxY + padding)
  return CGRect(
    x: originX,
    y: originY,
    width: maxBoundX - originX + 1,
    height: maxBoundY - originY + 1
  )
}

func integralCropRect(_ rect: CGRect, maxWidth: Int, maxHeight: Int) -> CGRect {
  let x = max(0, Int(floor(rect.origin.x)))
  let y = max(0, Int(floor(rect.origin.y)))
  let width = min(maxWidth - x, Int(ceil(rect.width)))
  let height = min(maxHeight - y, Int(ceil(rect.height)))
  return CGRect(x: x, y: y, width: max(1, width), height: max(1, height))
}

func aspectFillSquareImage(from image: CGImage, cropRect: CGRect, canvasSize: Int) throws -> CGImage {
  let crop = integralCropRect(cropRect, maxWidth: image.width, maxHeight: image.height)
  guard let cropped = image.cropping(to: crop) else {
    throw NSError(domain: "AssetGen", code: 3, userInfo: [NSLocalizedDescriptionKey: "Unable to crop image"])
  }

  let colorSpace = CGColorSpaceCreateDeviceRGB()
  guard
    let context = CGContext(
      data: nil,
      width: canvasSize,
      height: canvasSize,
      bitsPerComponent: 8,
      bytesPerRow: canvasSize * 4,
      space: colorSpace,
      bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue
    )
  else {
    throw NSError(domain: "AssetGen", code: 4, userInfo: [NSLocalizedDescriptionKey: "Unable to create output context"])
  }

  context.interpolationQuality = .high
  let scale = max(CGFloat(canvasSize) / CGFloat(cropped.width), CGFloat(canvasSize) / CGFloat(cropped.height))
  let drawWidth = CGFloat(cropped.width) * scale
  let drawHeight = CGFloat(cropped.height) * scale
  let drawRect = CGRect(
    x: (CGFloat(canvasSize) - drawWidth) / 2,
    y: (CGFloat(canvasSize) - drawHeight) / 2,
    width: drawWidth,
    height: drawHeight
  )
  context.draw(cropped, in: drawRect)

  guard let output = context.makeImage() else {
    throw NSError(domain: "AssetGen", code: 5, userInfo: [NSLocalizedDescriptionKey: "Unable to render output image"])
  }
  return output
}

func resizedSquareImage(from image: CGImage, size: Int) throws -> CGImage {
  let colorSpace = CGColorSpaceCreateDeviceRGB()
  guard
    let context = CGContext(
      data: nil,
      width: size,
      height: size,
      bitsPerComponent: 8,
      bytesPerRow: size * 4,
      space: colorSpace,
      bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue
    )
  else {
    throw NSError(domain: "AssetGen", code: 6, userInfo: [NSLocalizedDescriptionKey: "Unable to create resize context"])
  }
  context.interpolationQuality = .high
  context.draw(image, in: CGRect(x: 0, y: 0, width: size, height: size))
  guard let output = context.makeImage() else {
    throw NSError(domain: "AssetGen", code: 7, userInfo: [NSLocalizedDescriptionKey: "Unable to resize image"])
  }
  return output
}

func writePNG(_ image: CGImage, to path: String) throws {
  let rep = NSBitmapImageRep(cgImage: image)
  guard let data = rep.representation(using: .png, properties: [:]) else {
    throw NSError(domain: "AssetGen", code: 8, userInfo: [NSLocalizedDescriptionKey: "Unable to create PNG data"])
  }
  try data.write(to: URL(fileURLWithPath: path))
}

func qrCodeImage(for string: String, size: Int) throws -> CGImage {
  let filter = CIFilter(name: "CIQRCodeGenerator")
  filter?.setValue(Data(string.utf8), forKey: "inputMessage")
  filter?.setValue("H", forKey: "inputCorrectionLevel")
  guard let outputImage = filter?.outputImage else {
    throw NSError(domain: "AssetGen", code: 9, userInfo: [NSLocalizedDescriptionKey: "Unable to generate QR image"])
  }

  let context = CIContext(options: nil)
  guard let qrImage = context.createCGImage(outputImage, from: outputImage.extent) else {
    throw NSError(domain: "AssetGen", code: 10, userInfo: [NSLocalizedDescriptionKey: "Unable to render QR image"])
  }
  let qrScale = min(CGFloat(size - 160) / CGFloat(qrImage.width), CGFloat(size - 160) / CGFloat(qrImage.height))
  let drawWidth = CGFloat(qrImage.width) * qrScale
  let drawHeight = CGFloat(qrImage.height) * qrScale

  let colorSpace = CGColorSpaceCreateDeviceRGB()
  guard
    let canvas = CGContext(
      data: nil,
      width: size,
      height: size,
      bitsPerComponent: 8,
      bytesPerRow: size * 4,
      space: colorSpace,
      bitmapInfo: CGImageAlphaInfo.premultipliedLast.rawValue
    )
  else {
    throw NSError(domain: "AssetGen", code: 11, userInfo: [NSLocalizedDescriptionKey: "Unable to create QR canvas"])
  }

  canvas.setFillColor(NSColor.white.cgColor)
  canvas.fill(CGRect(x: 0, y: 0, width: size, height: size))
  let drawRect = CGRect(
    x: (CGFloat(size) - drawWidth) / 2,
    y: (CGFloat(size) - drawHeight) / 2,
    width: drawWidth,
    height: drawHeight
  )
  canvas.interpolationQuality = .none
  canvas.draw(qrImage, in: drawRect)

  guard let finalImage = canvas.makeImage() else {
    throw NSError(domain: "AssetGen", code: 12, userInfo: [NSLocalizedDescriptionKey: "Unable to finalize QR image"])
  }
  return finalImage
}

let repoRoot = "/Users/Kenmo/mayumi_bijiris"
let sourcePath = repoRoot + "/customer-app/icons/icon-512.png"
let icon192Path = repoRoot + "/customer-app/icons/icon-192.png"
let icon512Path = repoRoot + "/customer-app/icons/icon-512.png"
let appleTouchPath = repoRoot + "/customer-app/icons/apple-touch-icon.png"
let qrPath = repoRoot + "/customer-app/install-qr.png"
let installUrl = "https://mayumi-josanin.github.io/mayumi_bijiris/customer-app/"

let sourceImage = try loadCGImage(at: sourcePath)
let rgba = try rgbaImage(from: sourceImage)
let contentBounds = nonWhiteBounds(in: rgba)
let masterIcon = try aspectFillSquareImage(from: sourceImage, cropRect: contentBounds, canvasSize: 512)
let icon192 = try resizedSquareImage(from: masterIcon, size: 192)
let appleTouch = try resizedSquareImage(from: masterIcon, size: 180)

try writePNG(masterIcon, to: icon512Path)
try writePNG(icon192, to: icon192Path)
try writePNG(appleTouch, to: appleTouchPath)

print("Updated customer app icons:")
print(icon512Path)
print(icon192Path)
print(appleTouchPath)

if let qrImage = try? qrCodeImage(for: installUrl, size: 1200) {
  try writePNG(qrImage, to: qrPath)
  print(qrPath)
} else {
  print("QR generation skipped by local renderer failure. Generate \(qrPath) via network fetch if needed.")
}
