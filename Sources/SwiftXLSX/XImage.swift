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
//  Created by Kostiantyn Bohonos on 6/15/22.
//

import Foundation
#if os(macOS)
import Cocoa
#else
import UIKit
#endif

public enum XImageType:String {
    case png  = "png"
    case jpeg = "jpeg"
    case jpg = "jpg"
}

#if os(macOS)
extension NSBitmapImageRep {
    var pngData: Data? { representation(using: .png, properties: [:]) }
}
extension Data {
    var bitmap: NSBitmapImageRep? { NSBitmapImageRep(data: self) }
}
extension NSImage {
    func pngData() -> Data? {
        tiffRepresentation?.bitmap?.pngData
    }
}
#endif


public class XImage{
    var Key = ""
    var data:Data?
    var type:XImageType = .png
    
    func config(with data:Data, Key:String){
        self.data = data
        self.Key = Key
    }
    
    init(with data:Data, Key:String){
        self.config(with: data, Key: Key)
    }
    
    
    
    public init?(with image:ImageClass) {
        guard let data = image.pngData() else {return nil}
        self.config(with: data, Key: "\(XCS.checksum(data: data))")
    }
    
    public init?(with image:ImageClass, Key:String) {
        guard let data = image.pngData() else {return nil}
        self.config(with: data, Key: Key)
    }
    
    @discardableResult
    func Write(toPath path:String) -> Bool{
        var url = URL(fileURLWithPath: path)
        url = url.appendingPathComponent(self.Key)
        url = url.appendingPathExtension(self.type.rawValue)
        do {
            try self.data?.write(to: url, options: [.atomic])
        } catch {
            print("Error write file \(url) : \(error)")
            return false
        }
        return true
    }
}

public struct XImageCell {
    public let key: String
    public let size: CGSize
    
    public init(key: String, size: CGSize) { 
        self.key = key
        self.size = size
    }
}

extension XImageCell:Equatable{
    public static func == (lhs: Self, rhs: Self) -> Bool{
        lhs.key == rhs.key && lhs.size.equalTo(rhs.size)
    }
}

public class XImages{
    public static var list:[String:XImage] = [:]
    
    public static func append(with ximg:XImage) -> String{
        if list[ximg.Key] == nil {
            list[ximg.Key] = ximg
        }
        return ximg.Key
    }
    
    public static  func append(with image:ImageClass) -> String?{
        guard let ximg = XImage(with: image) else {return nil}
        return append(with: ximg)
    }
    
    public static  func append(with image:ImageClass, Key:String) -> String?{
        if list[Key] == nil {
            guard let ximg = XImage(with: image,Key: Key) else {return nil}
            return append(with: ximg)
        }else{
            return Key
        }
    }
    
    public static func removeAll(){
        self.list.removeAll()
    }
}
