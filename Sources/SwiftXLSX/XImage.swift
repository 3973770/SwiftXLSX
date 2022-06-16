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
import UIKit

public enum XImageType:String {
    case png  = "png"
    case jpeg = "jpeg"
    case jpg = "jpg"
}




public class XImage{
    var id = ""
    var dataicon:Data?
    var type:XImageType = .png
    
    init?(with image:UIImage) {
        guard let data = image.pngData() else {return nil}
        dataicon = data
        id = "\(XCS.checksum(data: dataicon!))"
    }
    
    init?(with ico:Xicons) {
        guard let data = ico.data else {return nil}
        dataicon = data
        id = ico.rawValue
    }
}

extension XImage: Equatable{
    public static func == (lhs: XImage, rhs: XImage) -> Bool{
        lhs.id == rhs.id
    }
}
