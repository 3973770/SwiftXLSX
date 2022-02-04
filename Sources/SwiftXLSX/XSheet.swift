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

import Foundation

public class XSheet{
    
    private static let ABC:[String] = ["A","B","C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"]
    public var title:String = ""
    private var cells:[XCell] = []
    var mergecells : [String] = []
    var merge : [XRect] = []
    var RowH : [Int:Int] = [:]
    var ColW : [Int:Int] = [:]
    private var indexcells : [String:XCell] = [:]
    var xml : String?
    var fix:XCoords = XCoords()
    
    public init() {}
    
    public init(_ title:String){
        self.title = title
    }
    
    public func ForRowSetHeight(_ row:Int,_ Height:Int){
        self.RowH[row] = Height
    }
    public func ForColumnSetWidth(_ column:Int,_ Width:Int){
        self.ColW[column] = Width
    }
    
    public var GetMaxRowCol : (row:Int,col:Int) {
        let res = XCoords()
        for Cell in self.cells {
            if res.row < Cell.coords!.row {
                res.row = Cell.coords!.row
            }
            if res.col < Cell.coords!.col {
                res.col = Cell.coords!.col
            }
        }
        return (res.row,res.col)
    }
    
    static func EncodeNumberABC(_ num:Int) -> String {
        var ret = ""
        if num < XSheet.ABC.count {
            ret = XSheet.ABC[num]
        }else{
            let whole:Int = num / XSheet.ABC.count
            let remain:Int = num - whole*XSheet.ABC.count
            
            ret = "\(XSheet.EncodeNumberABC(whole-1))\(XSheet.EncodeNumberABC(remain))"
        }
        return ret
    }
    
    public func MergeRect(_ rect:XRect) {
        let xcord = XSheet.EncodeNumberABC(rect.col-1)
        let ycord = "\(rect.row)"
        let x2cord = XSheet.EncodeNumberABC(rect.col+rect.width-2)
        let y2cord = "\(rect.row+rect.height-1)"
        self.merge.append(rect)
        self.mergecells.append("\(xcord)\(ycord):\(x2cord)\(y2cord)")
        
        if let maincell = self.Get(XCoords(row: rect.row, col: rect.col)) {
            for row in rect.row...rect.row+rect.height-1 {
                for col in rect.col...rect.col+rect.width-1 {
                    if col == rect.col, row == rect.row {
                        continue
                    }
                    
                    let cord = XCoords(row: row, col: col)
                    if self.Get(cord) == nil {
                        let cellnew = self.AddCell(cord)
                        cellnew.value = .text("")
                        cellnew.Border = maincell.Border
                        cellnew.Font = maincell.Font
                        cellnew.color = maincell.color
                        cellnew.colorbackground = maincell.colorbackground
                    }
                }
            }
        }
    }
    
    /// add cell with coords to current sheet
    public func AddCell(_ coords:XCoords) -> XCell{
        if let cell = self.Get(coords) {
            return cell
        }else{
            let cellnew = XCell(coords)
            self.cells.append(cellnew)
            AddCellToIndex(cellnew)
            return cellnew
        }
    }
    
    /// append exestit cell to current sheet
    func append(_ newElement: XCell){
        self.cells.append(newElement)
        AddCellToIndex(newElement)
    }
    
    func AddCellToIndex(_ cell:XCell){
        if let _ = self.indexcells[cell.coords!.address] {} else{
            self.indexcells[cell.coords!.address] = cell
        }
    }
    
    /// build indez of all cell in sheet
    public func buildindex(){
        self.indexcells.removeAll()
        for cell in self.cells {
            self.indexcells[cell.coords!.address] = cell
        }
    }
    
    /// get cell from sheer by coords
    public func Get(_ coords:XCoords) -> XCell? {
        guard let ret = self.indexcells[coords.address] else { return nil }
        return ret
    }
    
    func GetMaxWidth(_ col:Int,_ numrows:Int) -> Int {
        var maxw=50;
        for row in 1...numrows {
            let cell = self.Get(XCoords(row: row, col: col))
            if cell != nil, !cell!.nocalculatewidth, maxw < cell!.width {
                maxw = cell!.width
            }
        }
        return maxw
    }
    
}

extension XSheet: Sequence {
    public func makeIterator() -> Array<XCell>.Iterator {
        return cells.makeIterator()
    }
    
}

fileprivate extension String{
    
    func XSheetTitle() -> String
    {
        let syms:[String] = [":","\\","/","?","*","[","]","     ","    ","   ","  "]
        var str = "\(self)"
        for sym in syms {
            str = str.replacingOccurrences(of: sym, with: " ", options: NSString.CompareOptions.literal, range: nil)
        }
        return str
    }
}
