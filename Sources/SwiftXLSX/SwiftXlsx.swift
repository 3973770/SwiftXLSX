// Copyright 2022
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
//
//  Created by Kostiantyn Bohonos on 1/19/22.
//

#if os(macOS)
    import Cocoa
    public typealias FontClass  = NSFont
    public typealias ColorClass = NSColor
    public typealias ImageClass = NSImage
#else
    import UIKit
    public typealias FontClass  = UIFont
    public typealias ColorClass = UIColor
    public typealias ImageClass = UIImage
#endif

public struct SwiftXLSX {
    static func test() -> Bool {
        XWorkBook.test()
    }
}
