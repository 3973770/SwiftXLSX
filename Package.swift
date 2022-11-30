// swift-tools-version:5.5
// The swift-tools-version declares the minimum version of Swift required to build this package.

import PackageDescription

let package = Package(
    name: "SwiftXLSX",
    platforms: [
        .iOS(.v10),
        .macOS(.v11),
    ],
    products: [
        .library(
            name: "SwiftXLSX",
            targets: ["SwiftXLSX"]),
    ],
//     dependencies: [
//         .package(name: "ZipArchive", url: "https://github.com/ZipArchive/ZipArchive.git", from: "2.3.0")
//     ],
    targets: [
        .target(
            name: "SwiftXLSX",
//             dependencies: ["ZipArchive"]),
        .testTarget(
            name: "SwiftXLSXTests",
//             dependencies: ["SwiftXLSX"]),
    ]
)
