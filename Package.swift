// swift-tools-version: 6.1
// The swift-tools-version declares the minimum version of Swift required to build this package.

// swift-tools-version: 5.9
import PackageDescription

let package = Package(
    name: "ClassicXLS",
    platforms: [
        .iOS(.v13), .macOS(.v11), .tvOS(.v13), .watchOS(.v6)
    ],
    products: [
        .library(name: "ClassicXLS", targets: ["ClassicXLS"])
    ],
    targets: [
        .target(
            name: "ClassicXLS",
//            resources: [.process("Resources")] // for test fixtures later
        ),
        .testTarget(
            name: "ClassicXLSTests",
            dependencies: ["ClassicXLS"],
            resources: [.process("Fixtures")]
        )
    ]
)
