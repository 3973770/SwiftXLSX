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
import UIKit
import ZipArchive



public class XWorkBook{
    private var Sheets:[XSheet] = []
    
    // for styles
    private var Fonts:[UInt64:(String,Int)] = [:]
    private var Fills:[String] = []
    private var colorsid:[String] = []
    private var Bgcolor:[String] = []
    private var xfs:[UInt64:(String,Int)] = [:]
    private var Borders:[String] = []
    
    private var vals:[String] = []
    private var valss:Set<String> = Set([])
    private var CHARSIZE:[UInt64:CGFloat] = [:]
    
    public init() {}
    
    var count:Int {
        return self.Sheets.count
    }
    
    func removeAll(){
        self.Sheets.removeAll()
    }
    
    func append(_ newElement: XSheet){
        self.Sheets.append(newElement)
    }
    
    /// create and return new sheet
    public func NewSheet(_ title:String) -> XSheet{
        let Sheet = XSheet(title.XSheetTitle())
        self.append(Sheet)
        return Sheet
    }
    /// clear all local data for building xlsx
    private func clearlocaldata(){
        self.Fonts.removeAll()
        self.colorsid.removeAll()
        self.Fills.removeAll()
        self.Borders.removeAll()
        self.xfs.removeAll()
        self.vals.removeAll()
        self.vals.append("")
        self.valss.removeAll()
        self.valss.insert("")
        self.CHARSIZE.removeAll()// =
        self.Bgcolor.removeAll()
    }
    
    private func findFont(_ cell:XCell){
        let font = cell.Font!
        
        
        var idval:UInt64 = cell.Font!.bold ? 1 : 0
        idval += cell.Font!.italic ? 2 : 0
        idval += cell.Font!.strike ? 3 : 0
        idval += cell.Font!.underline ? 15 : 0
        idval += UInt64(cell.Font!.FontSize) * 10
        idval += (UInt64(cell.Font!.Font.ind())+1) * 10000

   
        let col:UIColor = cell.color
        if let hex = col.Hex {
            if let index = self.colorsid.firstIndex(of: hex) {
                idval += (UInt64(index)+1) * 1000000
            }else{
                self.colorsid.append(hex)
                idval += (UInt64(self.colorsid.count)+1) * 1000000
            }
        }
        
        if let (_,ind) = self.Fonts[idval] {
            cell.idFont = ind
        }else{
            let xml = "<font>\(font.bold ? "<b/>" : "")\(font.italic ? "<i/>" : "")\(font.strike ? "<strike/>" : "")\(font.underline ? "<u/>" : "")<sz val=\"\(font.FontSize)\"/><color rgb=\"\(cell.color.Hex!)\"/><name val=\"\(font.Font)\"/></font>"

            cell.idFont = self.Fonts.count
            self.Fonts[idval] = (xml,self.Fonts.count)
        }
    }
    
    private func findFills(_ cell:XCell){
        let hexcolor = cell.colorbackground.Hex!
        
        if let index = self.Bgcolor.firstIndex(of: hexcolor) {
            cell.idFill = index
        }else{
            self.Bgcolor.append(hexcolor)
            let Fontxml = "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"\(hexcolor)\"/><bgColor indexed=\"64\"/></patternFill></fill>"
            
            if let indexfill = self.Fills.firstIndex(of: Fontxml) {
                cell.idFill = indexfill
            }else{
                self.Fills.append(Fontxml)
                cell.idFill = self.Fills.count-1
            }
        }
    }
    
    private func findxf(_ cell:XCell,_ sheet:XSheet){
        cell.idStyle = -1
        var idval: UInt64 = cell.alignmentHorizontal.id()
        idval += cell.alignmentVertical.id() * 10
        idval += cell.Border ? 100 : 0
        idval += UInt64(cell.idFill) * 1000
        idval += UInt64(cell.idFont) * 1000000
        
        for merge in sheet.merge {
            if merge.inrect(cell.coords!){
                if let cellmain = sheet.Get(XCoords(row: merge.row, col: merge.col)) {
                    cell.idStyle = cellmain.idStyle
                    cell.idFont = cellmain.idFont
                    cell.idFill = cellmain.idFill
                    break
                }
            }
        }
        if cell.idStyle == -1 {
            if let (_,ind) = self.xfs[idval] {
                cell.idStyle = ind
            }else{
                let xf = "<xf fontId=\"\(cell.idFont)\" numFmtId=\"0\" fillId=\"\(cell.idFill)\" borderId=\"\(cell.Border ? "1" : "0")\" applyFont=\"1\" applyNumberFormat=\"0\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"\(cell.alignmentHorizontal.str())\" vertical=\"\(cell.alignmentVertical.str())\" textRotation=\"0\" wrapText=\"true\" shrinkToFit=\"false\"/></xf>"
                cell.idStyle = self.xfs.count
                self.xfs[idval] = (xf,self.xfs.count)
            }
        }
    }
    
    private func findVals(_ cell:XCell){
        cell.idVal = nil
        if let val = cell.value {
            switch val {
            case .text(let strval):
                let key:UInt64 =  cell.Font!.Font.ind()*1000000+UInt64(cell.Font!.FontSize)
                var sf:CGFloat = 0
                if let sizefont = self.CHARSIZE[key] {
                    sf = sizefont
                }else{
                    let size = cell.Font!.getfont!.Rectfor("0")
                    self.CHARSIZE[key] = size.width
                    sf = size.width
                }
                cell.width = Int(sf * CGFloat(strval.count) + 10)
                
                if self.valss.contains(strval) {
                    if let indexval = self.vals.firstIndex(of: strval) {
                        cell.idVal = indexval
                    }
                }else{
                    self.vals.append(strval)
                    self.valss.insert(strval)
                    cell.idVal = self.vals.count - 1
                }
            default:
                break
            }
        }
    }
    
    private func BuildStyles(){
        self.clearlocaldata()
        
        self.Fonts[0] = ("<font><sz val=\"10\"/><color rgb=\"FF000000\"/><name val=\"TrebuchetMS\"/></font>",0)
        
        self.Fills.append("<fill><patternFill patternType=\"solid\"><fgColor rgb=\"FFFFFFFF\"/><bgColor indexed=\"64\"/></patternFill></fill>")

        self.Bgcolor.append("FFFFFFFF")
        
        self.xfs[0] = ("<xf fontId=\"0\" numFmtId=\"0\" fillId=\"0\" borderId=\"0\" applyFont=\"1\" applyNumberFormat=\"0\" applyFill=\"1\" applyBorder=\"1\" applyAlignment=\"1\"><alignment horizontal=\"center\" vertical=\"center\" textRotation=\"0\" wrapText=\"true\" shrinkToFit=\"false\"/></xf>",0)
        
        self.Borders.append("<border/>");
        self.Borders.append("<border><left style=\"thin\"><color rgb=\"FF000000\"/></left><right style=\"thin\"><color rgb=\"FF000000\"/></right><top style=\"thin\"><color rgb=\"FF000000\"/></top><bottom style=\"thin\"><color rgb=\"FF000000\"/></bottom></border>");
        
        
        for sheet in self {
            for cell in sheet {
                self.findFont(cell)
                self.findFills(cell)
                
                self.findxf(cell , sheet)
                self.findVals(cell)
            }
        }
    }
    
    private func BuildSheet(_ sheet:XSheet){
        sheet.buildindex()
        let Xml:NSMutableString = NSMutableString()
        Xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
        Xml.append("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:mv=\"urn:schemas-microsoft-com:mac:vml\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">")
        
        let (numrows,numcols) = sheet.GetMaxRowCol
        if numrows>0 {
            Xml.append("<dimension ref=\"A1:\(XSheet.EncodeNumberABC(numcols-1))\(numrows)\"/>")
        }
        Xml.append("<sheetViews>")
        if (sheet.fix.row + sheet.fix.col) == 0 {
            Xml.append("<sheetView workbookViewId=\"0\" />")
        }else if sheet.fix.row > 0 , sheet.fix.col == 0 {
            Xml.append("<sheetView workbookViewId=\"0\">")
            Xml.append("<pane ySplit=\"\(sheet.fix.row)\" topLeftCell=\"A\(sheet.fix.row+1)\" activePane=\"bottomLeft\" state=\"frozen\"/>")
            Xml.append("</sheetView>")
        }else if sheet.fix.row == 0 , sheet.fix.col > 0 {
            Xml.append("<sheetView workbookViewId=\"0\">")
            Xml.append("<pane xSplit=\"\(sheet.fix.col)\" topLeftCell=\"\(XSheet.EncodeNumberABC(sheet.fix.col))1\" activePane=\"topRight\" state=\"frozen\"/>")
            Xml.append("</sheetView>")
        }else{
            Xml.append("<sheetView workbookViewId=\"0\">")
            Xml.append("<pane xSplit=\"\(sheet.fix.col)\" ySplit=\"\(sheet.fix.row)\" topLeftCell=\"\(XSheet.EncodeNumberABC(sheet.fix.col))\(sheet.fix.row+1)\" activePane=\"bottomRight\" state=\"frozen\"/>")
            Xml.append("</sheetView>")
        }
        Xml.append("</sheetViews>")
        
        Xml.append(" <sheetFormatPr defaultColWidth=\"8.625\" defaultRowHeight=\"14.4\" customHeight=\"1\" outlineLevelRow=\"0\" outlineLevelCol=\"0\"/>")
        Xml.append("<cols>")
       // Xml.append("<col min=\"1\" max=\"\(numcols)\" width=\"20\" customWidth=\"true\" style=\"0\"/>")
        
        for col:Int in 1...numcols {
            var wcalc:CGFloat = 50
            if sheet.ColW.count > 0 {
                let w = sheet.ColW[col]
                if w != nil {
                    wcalc = CGFloat(w!)
                }else{
                    wcalc  = CGFloat(sheet.GetMaxWidth(col, numrows))
                }
            }else{
                wcalc  = CGFloat(sheet.GetMaxWidth(col, numrows))
            }
            wcalc = wcalc * (0.16666016 * 1.2)//0.12499512
            
            Xml.append("<col min=\"\(col)\" max=\"\(col == numcols ? 16384 : col)\" width=\"\(wcalc)\" customWidth=\"1\" style=\"0\"/>")
        }
        Xml.append("</cols>")
        Xml.append("<sheetData>")
        if numrows == 0 {
            Xml.append("<row r=\"1\" spans=\"1:1\"><c r=\"A1\"/></row>")
        }else{
            for row:Int in 1...numrows {
                let colls:NSMutableString = NSMutableString()
                for col:Int in 1...numcols {
                    if let cell = sheet.Get(XCoords(row: row, col: col)) {
                        /// output  values
                        if let value = cell.value {
                            switch value {
                            case .text(_):
                                if cell.idVal != nil {
                                    /// this is text value and insert like id
                                    colls.append("<c r=\"\(XSheet.EncodeNumberABC(col-1))\(row)\" s=\"\(cell.idStyle)\" t=\"s\"><v>\(cell.idVal!)</v></c>")
                                }else{
                                    /// output empty cell
                                    colls.append("<c r=\"\(XSheet.EncodeNumberABC(col-1))\(row)\" s=\"\(cell.idStyle >= 0 ? cell.idStyle : 0)\" />")
                                }
                            case .double(let val):
                                colls.append("<c r=\"\(XSheet.EncodeNumberABC(col-1))\(row)\" s=\"\(cell.idStyle)\" ><v>\(String(format: "%.3f", val))</v></c>")
                            case .float(let val):
                                colls.append("<c r=\"\(XSheet.EncodeNumberABC(col-1))\(row)\" s=\"\(cell.idStyle)\" ><v>\(String(format: "%.3f", val))</v></c>")
                            case.integer(let val):
                                colls.append("<c r=\"\(XSheet.EncodeNumberABC(col-1))\(row)\" s=\"\(cell.idStyle)\" ><v>\(val)</v></c>")
                            case .long(let val):
                                colls.append("<c r=\"\(XSheet.EncodeNumberABC(col-1))\(row)\" s=\"\(cell.idStyle)\" ><v>\(val)</v></c>")
                            }
                        }else{
                            colls.append("<c r=\"\(XSheet.EncodeNumberABC(col-1))\(row)\" s=\"\(cell.idStyle >= 0 ? cell.idStyle : 0)\" />")
                        }
                        
                    }else{
                        /// no cell with coords (row,col)
                        colls.append("<c r=\"\(XSheet.EncodeNumberABC(col-1))\(row)\" s=\"0\" />")
                    }
                }
                if let ht = sheet.RowH[row] {
                    Xml.append("<row r=\"\(row)\" spans=\"1:\(numcols)\" ht=\"\(ht)\" customHeight=\"1\">")
                }else{
                    Xml.append("<row r=\"\(row)\" spans=\"1:\(numcols)\">")
                }
                Xml.append(String(colls))
                colls.setString("")
                Xml.append("</row>")
            }
        }
        Xml.append("</sheetData>")
        
        if sheet.mergecells.count > 0 {
            Xml.append("<mergeCells count=\"\(sheet.mergecells.count)\">")
            for addr in sheet.mergecells {
                Xml.append("<mergeCell ref=\"\(addr)\"/>")
            }
            Xml.append("</mergeCells>")
        }
        Xml.append("</worksheet>")
        sheet.xml = String(Xml)
        Xml.setString("")
    }
    
    
    private func BuildSheets() {
        for sheet in self {
            self.BuildSheet(sheet)
        }
    }
    
    
    /// Check exist directory and create if don't exist
    @discardableResult
    private func CheckCreateDirectory(path pathdirectory:String) -> Bool{
        let filemanager = FileManager.default
        if !filemanager.fileExists(atPath: pathdirectory) {
            do {
                try filemanager.createDirectory(atPath: pathdirectory, withIntermediateDirectories: true, attributes: nil)
            } catch {
                print("Error creating directory \(pathdirectory) : \(error.localizedDescription)")
                return false
            }
        }
        return true
    }
    
    /// remove file or directory
    @discardableResult
    private func RemoveFile(path pathfile:String) -> Bool{
        do {
            try FileManager.default.removeItem(atPath: pathfile)
        } catch {
            print("Error remove file \(pathfile) : \(error.localizedDescription)")
            return false
        }
        return true
    }
    
    /// write string buffer to file
    @discardableResult
    private func Write(data strData:String, tofile path:String) -> Bool{
        do {
            try strData.write(toFile: path, atomically: true, encoding: .utf8)
        } catch {
            print("Error write file \(path) : \(error.localizedDescription)")
            return false
        }
        return true
    }
    
    private var rels:String {
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"
    }
    
    private var SharedStrings:String {
        let Xml:NSMutableString = NSMutableString()
        Xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
        Xml.append("<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" uniqueCount=\"\(self.vals.count)\">")
        for val in vals {
            Xml.append("<si><t>\(val.XmlPrep())</t></si>")
        }
        Xml.append("</sst>")
        return String(Xml)
    }
    
    private var StyleStrings:String {
        let Xml:NSMutableString = NSMutableString()
        
        Xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
        Xml.append("<styleSheet xml:space=\"preserve\" xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">")
        Xml.append("<numFmts count=\"0\"/>")
        
        if Fonts.isEmpty {
            Xml.append("<fonts count=\"0\"/>")
        }else{
            Xml.append("<fonts count=\"\(Fonts.count)\">")
            let ar = Fonts.values.sorted(by: {  $0.1 < $1.1})
            for (font,_) in ar {
                Xml.append(font)
            }
            Xml.append("</fonts>")
        }
        
        if Fills.isEmpty {
            Xml.append("<fills count=\"0\"/>")
        }else{
            Xml.append("<fills count=\"\(Fills.count)\">")
            for fill in Fills {
                Xml.append(fill)
            }
            Xml.append("</fills>")
        }
        
        if Borders.isEmpty {
            Xml.append("<borders count=\"0\"/>")
        }else{
            Xml.append("<borders count=\"\(Borders.count)\">")
            for border in Borders {
                Xml.append(border)
            }
            Xml.append("</borders>")
        }
       
        Xml.append("<cellStyleXfs count=\"1\">")
        Xml.append("<xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/>")
        Xml.append("</cellStyleXfs>")
        
        if xfs.isEmpty {
            Xml.append("<cellXfs count=\"0\"/>")
        }else{
            Xml.append("<cellXfs count=\"\(xfs.count)\">")
            
            let ar = xfs.values.sorted(by: {  $0.1 < $1.1})
            for (xf,_) in ar {
                Xml.append(xf)
            }
            Xml.append("</cellXfs>")
        }

        Xml.append("<cellStyles count=\"1\">")
        Xml.append("<cellStyle name=\"Normal\" xfId=\"0\" builtinId=\"0\"/></cellStyles>")
        Xml.append("<dxfs count=\"0\"/>")
        Xml.append("</styleSheet>")

        return String(Xml)
    }
    
    private var ContentTypesStrings:String {
        let Xml:NSMutableString = NSMutableString()
        Xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
        Xml.append("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">")
        Xml.append("<Default ContentType=\"application/xml\" Extension=\"xml\"/>")
        Xml.append("<Default ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" Extension=\"rels\"/>")
        for i in 1...Sheets.count {
            Xml.append("<Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" PartName=\"/xl/worksheets/sheet\(i).xml\"/>")
        }
        
        Xml.append("<Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\" PartName=\"/xl/sharedStrings.xml\"/>")
        Xml.append("<Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" PartName=\"/xl/styles.xml\"/>")
        Xml.append("<Override ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" PartName=\"/xl/workbook.xml\"/>")
        Xml.append("</Types>")
        return String(Xml)
    }
    
    private var WorkBookXmlRelsStrings:String {
        let Xml:NSMutableString = NSMutableString()
        Xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
        Xml.append("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">")
        Xml.append("<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\"/>")
        Xml.append("<Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"sharedStrings.xml\"/>")
        for i in 1...Sheets.count {
            Xml.append("<Relationship Id=\"rId\(i+2)\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet\(i).xml\"/>")
        }
        Xml.append("</Relationships>")
        return String(Xml)
    }
    
    private var WorkBookXmlStrings:String {
        let Xml:NSMutableString = NSMutableString()
        
        Xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n")
        Xml.append("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:mv=\"urn:schemas-microsoft-com:mac:vml\" xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\" xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">")
        Xml.append("<workbookPr/>")
        Xml.append("<sheets>")
        for i in 1...Sheets.count {
            Xml.append("<sheet state=\"visible\" name=\"\(Sheets[i-1].title)\" sheetId=\"\(i)\" r:id=\"rId\(i+2)\"/>")
        }
        Xml.append("</sheets>")
        Xml.append("<definedNames/>")
        Xml.append("<calcPr/>")
        Xml.append("</workbook>")
        return String(Xml)
    }
    
    private var HeadSheedXML: String {
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>"
    }
    
    /// prepare files for Xlsx file
    private func preparefiles(for filename:String) -> String{
        var CachePath = NSSearchPathForDirectoriesInDomains(.cachesDirectory, .userDomainMask, true)[0]
        
        CachePath = "\(CachePath)/tmpxls"
        self.CheckCreateDirectory(path: CachePath)
        
        let FolderId = "folderxls\(arc4random())"
        let BasePath = "\(CachePath)/\(FolderId)"
        self.CheckCreateDirectory(path: BasePath)
        self.CheckCreateDirectory(path: "\(BasePath)/_rels")
        self.CheckCreateDirectory(path: "\(BasePath)/xl")
        self.CheckCreateDirectory(path: "\(BasePath)/xl/_rels")
        self.CheckCreateDirectory(path: "\(BasePath)/xl/worksheets")
        self.CheckCreateDirectory(path: "\(BasePath)/xl/worksheets/_rels")
        
        
        self.Write(data: self.rels, tofile: "\(BasePath)/_rels/.rels")
        self.Write(data: self.SharedStrings, tofile: "\(BasePath)/xl/sharedStrings.xml")
        self.Write(data: self.StyleStrings, tofile: "\(BasePath)/xl/styles.xml")
        self.Write(data: self.ContentTypesStrings, tofile: "\(BasePath)/[Content_Types].xml")
        self.Write(data: self.WorkBookXmlRelsStrings, tofile: "\(BasePath)/xl/_rels/workbook.xml.rels")
        self.Write(data: self.WorkBookXmlStrings, tofile: "\(BasePath)/xl/workbook.xml")
        
        var i = 1
        for sheet in Sheets {
            self.Write(data: self.HeadSheedXML, tofile: "\(BasePath)/xl/worksheets/_rels/sheet\(i).xml.rels")
            self.Write(data: sheet.xml!, tofile: "\(BasePath)/xl/worksheets/sheet\(i).xml")
            i += 1
        }

        
        let filepath = "\(CachePath)/\(filename)"
        _ = SSZipArchive.createZipFile(atPath: filepath,
                                                 withContentsOfDirectory: BasePath,
                                                 keepParentDirectory: false,
                                                 compressionLevel: -1,
                                                 password: nil,
                                                 aes: true,
                                                 progressHandler: nil)
        self.RemoveFile(path: BasePath)
        return filepath
    }
    
    
    /// write xlxs file and return path
    public func save(_ filename:String) -> String {
        self.BuildStyles()
        self.BuildSheets()
        return self.preparefiles(for: filename)
    }
    
    /// generate example xlsx file
    static func test() -> Bool {
   
           let book = XWorkBook()
   
           let color:[UIColor] = [.darkGray, .green, .lightGray, .orange, .systemPink, .cyan, .purple, .magenta, .blue]
           let colortext:[UIColor] = [.darkGray, .black, .white, .darkText, .lightText]
   
   
           func GetRandomFont() -> XFontName {
               let cases = XFontName.allCases
               return cases[Int.random(in: 0..<cases.count)]
           }
   
   
           var sheet = book.NewSheet("Invoice")
   
           var cell = sheet.AddCell(XCoords(row: 2, col: 6))
           cell.Cols(txt: .white, bg: .systemOrange)
           cell.value = .text("INVOICE")
           cell.Font = XFont(.TrebuchetMS, 16,true)
           cell.alignmentHorizontal = .center
   
           cell = sheet.AddCell(XCoords(row: 3, col: 6))
           cell.Cols(txt: .white, bg: .systemOrange)
           cell.value = .text("#12345")
           cell.Font = XFont(.TrebuchetMS, 12,true)
           cell.alignmentHorizontal = .left
   
           cell = sheet.AddCell(XCoords(row: 2, col: 1))
           cell.value = .text("Your company name")
           cell.Font = XFont(.TrebuchetMS, 16,true)
   
           cell = sheet.AddCell(XCoords(row: 3, col: 1))
           cell.value = .text("[Address Line 1]")
   
           cell = sheet.AddCell(XCoords(row: 4, col: 1))
           cell.value = .text("[Address Line 2]")
   
           cell = sheet.AddCell(XCoords(row: 5, col: 1))
           cell.value = .text("[Address Line 3]")
   
           cell = sheet.AddCell(XCoords(row: 7, col: 1))
           cell.color = .systemOrange
           cell.value = .text("Bill To:")
           cell.Font = XFont(.TrebuchetMS, 12,true)
   
           cell = sheet.AddCell(XCoords(row: 8, col: 1))
           cell.value = .text("[Address Line 1]")
   
           cell = sheet.AddCell(XCoords(row: 9, col: 1))
           cell.value = .text("[Address Line 2]")
   
           cell = sheet.AddCell(XCoords(row: 10, col: 1))
           cell.value = .text("[Address Line 3]")
   
           /// date
           cell = sheet.AddCell(XCoords(row: 13, col: 1))
           cell.color = .systemOrange
           cell.value = .text("Invoice Date")
           cell.Font = XFont(.TrebuchetMS, 12,true)
   
           cell = sheet.AddCell(XCoords(row: 14, col: 1))
           cell.value = .text("01/22/2022")
   
           /// term
           cell = sheet.AddCell(XCoords(row: 13, col: 2))
   
           cell.value = .text("Terms")
           cell.Font = XFont(.TrebuchetMS, 12,true)
   
           cell = sheet.AddCell(XCoords(row: 14, col: 2))
           cell.value = .text("30 days")
   
           /// Due Date
           cell = sheet.AddCell(XCoords(row: 13, col: 3))
           cell.color = .systemOrange
           cell.value = .text("Due Date")
           cell.Font = XFont(.TrebuchetMS, 12,true)
   
           cell = sheet.AddCell(XCoords(row: 14, col: 3))
           cell.value = .text("02/20/2022")
   
           /// table
           /// title
           var line = 16
           cell = sheet.AddCell(XCoords(row: line, col: 1))
           cell.Cols(txt: .white, bg: .systemOrange)
           cell.Border = true
           cell.value = .text("Description")
   
           cell = sheet.AddCell(XCoords(row: line, col: 4))
           cell.Cols(txt: .white, bg: .systemOrange)
           cell.Border = true
           cell.value = .text("Qty")
           cell.Font = XFont(.TrebuchetMS, 10,true)
   
           cell = sheet.AddCell(XCoords(row: line, col: 5))
           cell.Cols(txt: .white, bg: .systemOrange)
           cell.Border = true
           cell.value = .text("Unit Price")
           cell.Font = XFont(.TrebuchetMS, 10,true)
   
           cell = sheet.AddCell(XCoords(row: line, col: 6))
           cell.Cols(txt: .white, bg: .systemOrange)
           cell.Border = true
           cell.value = .text("Amount")
           cell.Font = XFont(.TrebuchetMS, 10,true)
   
           /// line
           line += 1
           cell = sheet.AddCell(XCoords(row: line, col: 1))
           cell.Border = true
           cell.value = .text("item #1")
   
           cell = sheet.AddCell(XCoords(row: line, col: 4))
           cell.Border = true
           cell.value = .double(3)
           cell.alignmentHorizontal = .right
   
           cell = sheet.AddCell(XCoords(row: line, col: 5))
           cell.Border = true
           cell.value = .double(50)
           cell.alignmentHorizontal = .right
   
           cell = sheet.AddCell(XCoords(row: line, col: 6))
           cell.Border = true
           cell.value = .double(150)
           cell.alignmentHorizontal = .right
   
           /// line
           line += 1
           cell = sheet.AddCell(XCoords(row: line, col: 1))
           if #available(iOS 13.0, *) {
               cell.Cols(txt: .black, bg: .systemGray6)
           } else {
               cell.Cols(txt: .black, bg: .lightGray)
           }
           cell.Border = true
           cell.value = .text("item #2")
   
           cell = sheet.AddCell(XCoords(row: line, col: 4))
           if #available(iOS 13.0, *) {
               cell.Cols(txt: .black, bg: .systemGray6)
           } else {
               cell.Cols(txt: .black, bg: .lightGray)
           }
           cell.Border = true
           cell.value = .double(2)
           cell.alignmentHorizontal = .right
   
           cell = sheet.AddCell(XCoords(row: line, col: 5))
           if #available(iOS 13.0, *) {
               cell.Cols(txt: .black, bg: .systemGray6)
           } else {
               cell.Cols(txt: .black, bg: .lightGray)
           }
           cell.Border = true
           cell.value = .double(100)
           cell.alignmentHorizontal = .right
   
           cell = sheet.AddCell(XCoords(row: line, col: 6))
           if #available(iOS 13.0, *) {
               cell.Cols(txt: .black, bg: .systemGray6)
           } else {
               cell.Cols(txt: .black, bg: .lightGray)
           }
           cell.Border = true
           cell.value = .double(200)
           cell.alignmentHorizontal = .right
   
           /// line
           line += 1
           cell = sheet.AddCell(XCoords(row: line, col: 1))
           cell.Border = true
           cell.value = .text("item #3")
   
           cell = sheet.AddCell(XCoords(row: line, col: 4))
           cell.Border = true
           cell.value = .double(4)
           cell.alignmentHorizontal = .right
   
           cell = sheet.AddCell(XCoords(row: line, col: 5))
           cell.Border = true
           cell.value = .double(200)
           cell.alignmentHorizontal = .right
   
           cell = sheet.AddCell(XCoords(row: line, col: 6))
           cell.Border = true
           cell.value = .double(800)
           cell.alignmentHorizontal = .right
   
           line += 2
   
           cell = sheet.AddCell(XCoords(row: line, col: 1))
           cell.value = .text("Thank you for your business!")
           cell.Font = XFont(.TrebuchetMS, 10, false, true)
           cell.alignmentHorizontal = .left
   
           cell = sheet.AddCell(XCoords(row: line, col: 5))
           cell.Cols(txt: .white, bg: .systemOrange)
           cell.value = .text("Total")
           cell.Font = XFont(.TrebuchetMS, 11, true)
           cell.alignmentHorizontal = .left
   
           cell = sheet.AddCell(XCoords(row: line, col: 6))
           cell.Cols(txt: .white, bg: .systemOrange)
           cell.value = .double(1100)
           cell.Font = XFont(.TrebuchetMS, 11, true)
           cell.alignmentHorizontal = .right
   
           line += 2
   
           cell = sheet.AddCell(XCoords(row: line, col: 1))
           cell.value = .text("Payment Options")
           cell.Font = XFont(.TrebuchetMS, 12, true, false)
           cell.alignmentHorizontal = .left
   
           line += 1
   
           cell = sheet.AddCell(XCoords(row: line, col: 1))
           cell.value = .text("Enter PayPal email address or bank account number here")
           cell.Font = XFont(.TrebuchetMS, 10, false, true)
           cell.alignmentHorizontal = .left
   
           sheet.buildindex()
   
           sheet.MergeRect(XRect(2, 1, 5, 1))
           sheet.MergeRect(XRect(3, 1, 5, 1))
           sheet.MergeRect(XRect(4, 1, 5, 1))
           sheet.MergeRect(XRect(5, 1, 5, 1))
   
           sheet.MergeRect(XRect(16, 1, 3, 1))
           sheet.MergeRect(XRect(17, 1, 3, 1))
           sheet.MergeRect(XRect(18, 1, 3, 1))
           sheet.MergeRect(XRect(19, 1, 3, 1))
   
           sheet.MergeRect(XRect(23, 1, 3, 1))
           sheet.MergeRect(XRect(24, 1, 3, 1))
   
   
           sheet.ForColumnSetWidth(1, 100)
           sheet.ForColumnSetWidth(2, 70)
           sheet.ForColumnSetWidth(3, 60)
           sheet.ForColumnSetWidth(4, 70)
           sheet.ForColumnSetWidth(5, 70)
           sheet.ForColumnSetWidth(6, 90)

   
           sheet = book.NewSheet("Perfomance1 Sheet")
           for col in 1...20 {
               sheet.ForColumnSetWidth(col,Int.random(in: 50..<100))
               for row in 1...1000 {
                   let cell = sheet.AddCell(XCoords(row: row, col: col))
                   cell.value = .integer(Int.random(in: 100..<200000))
                   cell.Font = XFont(GetRandomFont(), Int.random(in: 10..<20))
                   cell.color = colortext[Int.random(in: 0..<colortext.count)]
                   cell.colorbackground = color[Int.random(in: 0..<color.count)]
                   cell.Border = true
                   cell.alignmentHorizontal = .center
                   cell.alignmentVertical = .center
               }
           }
   
           sheet = book.NewSheet("Perfomance2 Sheet")
           for col in 1...20 {
               sheet.ForColumnSetWidth(col,Int.random(in: 50..<100))
               for row in 1...1000 {
                   let cell = sheet.AddCell(XCoords(row: row, col: col))
                   cell.value = .text("\(row):\(col)")
                   cell.Font = XFont(GetRandomFont(), Int.random(in: 10..<20))
                   cell.color = colortext[Int.random(in: 0..<colortext.count)]
                   cell.colorbackground = color[Int.random(in: 0..<color.count)]
                   cell.Border = true
                   cell.alignmentHorizontal = .center
                   cell.alignmentVertical = .center
               }
           }
   
           let fileid = book.save("example.xlsx")
           print("<<<File XLSX generated!>>>")
           print("\(fileid)")
           return true
       }
}

extension XWorkBook: Sequence {
    public func makeIterator() -> Array<XSheet>.Iterator {
        return Sheets.makeIterator()
    }
}

fileprivate extension UIFont {
    /// Calculate size text for current font
    func Rectfor(_ str:String)-> CGSize {
        let fontAttributes = [NSAttributedString.Key.font: self]
        let size = (str as NSString).size(withAttributes: fontAttributes)
        return size
    }
}

fileprivate extension String{
    func XmlPrep() -> String
    {
        let sm = NSMutableString()
        for (_, char) in self.enumerated() {
            switch char {
            case "&":
                sm.append("&amp;")
            case "\"":
                sm.append("&quot;")
            case "'":
                sm.append("&#x27;")
            case ">":
                sm.append("&gt;")
                case "<":
                   sm.append("&lt;")
                default:
                   sm.append(String(char))
            }
        }
        return String(sm)
    }
    
    func XSheetTitle() -> String
    {
        let sm = NSMutableString()
        for (_, char) in self.enumerated() {
            switch char {
            case ":":
                sm.append(" ")
            case "\\":
                sm.append(" ")
            case "\"":
                sm.append(" ")
            case "?":
                sm.append(" ")
            case "*":
                sm.append(" ")
                case "[":
                   sm.append(" ")
            case "]":
               sm.append(" ")
                default:
                   sm.append(String(char))
            }
        }
        return String(sm)
    }
}

extension UIColor {
    static var hexdict:[UInt64:String] = [:]
    /// encode color to HEX format AARRGGBB use cache for optimization
    var Hex:String?{
        var r: CGFloat = 0
        var g: CGFloat = 0
        var b: CGFloat = 0
        var a: CGFloat = 0
        self.getRed(&r, green: &g, blue: &b, alpha: &a)
        let ri = lroundf(Float(r) * 255)
        let gi = lroundf(Float(g) * 255)
        let bi = lroundf(Float(b) * 255)
        let ai = lroundf(Float(a) * 255)
        
        let idcolor = UInt64(ai)*1000000000+UInt64(ri)*1000000+UInt64(gi)*1000+UInt64(bi)
        if let hexcol = UIColor.hexdict[idcolor] {
            return hexcol
        }else{
            let hexcolgen = String(format: "%02lX%02lX%02lX%02lX",ai,ri,gi,bi)
            UIColor.hexdict[idcolor] = hexcolgen
            return hexcolgen
        }
    }
}
