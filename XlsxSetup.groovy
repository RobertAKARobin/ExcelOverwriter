package org.ontel

import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.*
import org.apache.poi.xssf.streaming.*

class XlsxSetup{
	
	XlsxSetup(){
				
		XSSFSheet.metaClass.findHeader{regex ->
			println "Looking for header matching $regex in sheet ${getSheetName()}..."
			def lastRowNumber = 5
			for(rowNumber in 0..lastRowNumber){
				println "Checking row $rowNumber..."
				def row = delegate.getRow(rowNumber)
				if(row == null) continue
				for(def cell in row){
					def columnName = cell?.getValue()
					if(columnName.toString().find(regex) != null){
						println "Found the header matching $regex!"
						return ["x": cell.getColumnIndex(), "y": rowNumber]
					}
				}
			}
			println "No header found in first $lastRowNumber rows."
			return false
		}
		
		XSSFSheet.metaClass.mapHeadersToIndexes{ headers, originY = 0 ->
			def outputMap = [:]
			def headerRow = delegate.getRow(originY)
			
			for(cell in headerRow){
				def cellValue = cell.getValue()
				if(!headers.contains(cellValue)) continue
				outputMap[cellValue] = cell.getColumnIndex()
			}
			println "Got all the headers with their indexes:"
			println outputMap
			return outputMap
		}
		
		XSSFSheet.metaClass.doByRow{ keyIndex, doSomething, startRowNum = 0, endRowNum = 0 ->
			if(endRowNum == 0) endRowNum = delegate.getPhysicalNumberOfRows()
			
			for(rowNum in startRowNum..endRowNum){
				def row = delegate.getRow(rowNum)
				if(row == null){
					println "Row $rowNum is null"
					break
				}
				
				def keyValue = row.getCell(keyIndex)?.getValue()
				if(keyValue == null || keyValue == ""){
					println "Row $rowNum has no key"
					continue
				}
				doSomething(row, keyValue)
				
			}
			println "Iterated over ${endRowNum - startRowNum} rows of data."
		}
		
		XSSFSheet.metaClass.getDataByHeaders{ headers, startCellNum = 0, startRowNum = 0 ->
			def output = [:]
			delegate.doByRow(startCellNum, { row, key->
				
				def rowData = [:]
				headers.each{ header, index->
					def data = row.getCell(index)?.getValue()
					rowData[header] = data
				}
				output[key] = rowData
				
			}, startRowNum)
			return output
		}
		
		XSSFSheet.metaClass.overwriteData{
			inputData,
			headerMap,
			startCellNum = 0,
			startRowNum = 0 ->
				delegate.doByRow(
					startCellNum, { row, rowKey ->
						for(header in headerMap){
							def cell	= row.getCell(header.value.outputIndex)
							def oldData	= cell?.getValue()
							
							if(oldData == rowKey) 			continue
							if(inputData[rowKey] == null) 	continue
							
							def newData	= inputData[rowKey][header.key]
							if(oldData == newData) 			continue
							cell.setCellValue(newData)
							println "Changing $rowKey $header.value.outputIndex from $oldData to $newData"
						}
					},
				startRowNum)
			;
		}
		
		XSSFCell.metaClass.cellTypes = [
		"numeric"	: 0,
		"string"	: 1,
		"formula"	: 2,
		"blank"		: 3,
		"boolean"	: 4,
		"error"		: 5
		]
		
		XSSFCell.metaClass.getValue{ -> 
			def output
			switch (delegate?.getCellType()){
				case delegate.cellTypes.numeric:
					output = delegate.getNumericCellValue()
					if(output % 1 == 0){
						output = output.toInteger()
					}else{
						output = output.toFloat()
					}
					break
				case delegate.cellTypes.formula:
					switch (delegate.getCachedFormulaResultType()){
						case delegate.cellTypes.string:
							output = delegate.getStringCellValue()
							break
						case delegate.cellTypes.numeric:
							output = delegate.getNumericCellValue()
							if(new DateUtil().isCellDateFormatted(delegate)){
								output = delegate.getDateCellValue()
							}
						default:
							output = ""
							break
					}
				case delegate.cellTypes.error:
					output = ""
					break
				case null:
					output = ""
					break
				default:
					output = delegate.getStringCellValue().toString()
					break
			}
			return output;
		}
						
	}
	
}
