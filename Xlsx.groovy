package org.ontel

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.xssf.streaming.*
import org.apache.poi.openxml4j.*
import org.apache.poi.openxml4j.opc.OPCPackage

class Xlsx{
	
	def filepath
	def startTime
	def openTime
	def closeTime
	def styles = [:]
	
	def inputFile
	def inputStream
	def outputStream
	def opc
	def workbook
	
		
	Xlsx(path){
		
		filepath	= path
		startTime	= System.currentTimeMillis()
		println "$startTime: $filepath opening..."
		inputFile	= new File(filepath)
		inputStream	= new FileInputStream(inputFile)
		opc 		= OPCPackage.open(inputStream)
		workbook	= WorkbookFactory.create(opc)
		openTime	= System.currentTimeMillis()
		println "$openTime: $filepath is open. It took ${(openTime - startTime)/1000} s."
	
	}
	
	public writeAndClose(filePathOut){
		println "Opening output stream..."
		outputStream = new FileOutputStream(new File(filePathOut));
		println "Writing to output stream..."
		workbook.write(outputStream);
		inputStream.close()
		outputStream.close()
		opc.close()
		
	}
	
	public close(what){
		closeTime = System.currentTimeMillis()
		println "$closeTime: $filepath is closed. It took ${(closeTime - openTime)/1000} s."
	}
	
	def findSheet(regex){
		println "Looking for sheet matching $regex in ${this.filepath}..."
		
		def sheet
		def numberOfSheets = workbook.getNumberOfSheets()
		for(sheetIndex in 0..numberOfSheets){
			sheet = workbook.getSheetAt(sheetIndex)
			def sheetName = sheet.getSheetName()
			println "Sheet $sheetIndex is named $sheetName..."
			if(sheetName.find(regex) != null){
				println "Found the sheet!"
				return sheet
			}
		}
		println "End of sheets. No match found."
		return false
	}
								
}
