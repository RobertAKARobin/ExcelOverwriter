package org.ontel

import java.nio.file.Files
import java.nio.file.Paths

class UpdateByHeaders{

	def inputFile
	def inputSheet
	def inputSheetOrigin
	def inputSheetHeaders = [:]
	def inputData = [:]
	
	def outputFile
	def outputSheet
	def outputSheetOrigin
	def outputSheetHeaders
	def outputData = [:]
	
	UpdateByHeaders(
		inputPath,
		inputSheetRegex,
		inputKeyRegex,
		outputPathIn,
		outputPathOut,
		outputSheetRegex,
		outputKeyRegex,
		headerMap
	){
		
		headerMap.each{ inputHeader, outputHeader->
			if(outputHeader == "same") headerMap[inputHeader] = inputHeader;
		}
				
		inputFile 			= new Xlsx(inputPath)
		inputSheet 			= inputFile.findSheet(inputSheetRegex)
		inputSheetOrigin	= inputSheet.findHeader(inputKeyRegex)
		inputSheetHeaders	= inputSheet.mapHeadersToIndexes(
								headerMap.keySet(),
								inputSheetOrigin.y
								)
		inputData 			= inputSheet.getDataByHeaders(
								inputSheetHeaders,
								inputSheetOrigin.x,
								inputSheetOrigin.y + 1
								)
		inputFile.close()
		
		outputFile			= new Xlsx(outputPathOut)
		outputSheet 		= outputFile.findSheet(outputSheetRegex)
		outputSheetOrigin	= outputSheet.findHeader(outputKeyRegex)
		outputSheetHeaders	= outputSheet.mapHeadersToIndexes(
								headerMap.values(),
								outputSheetOrigin.y
								)
		
		headerMap.each{ inputHeader, outputHeader->
			headerMap[inputHeader] = [
					"outputHeader"	: outputHeader,
					"inputIndex"	: inputSheetHeaders[inputHeader],
					"outputIndex"	: outputSheetHeaders[outputHeader]
			]
		}
						
		outputSheet.overwriteData(
								inputData,
								headerMap,
								outputSheetOrigin.x,
								outputSheetOrigin.y + 1
		)
		outputFile.writeAndClose(outputPathOut)
	}
	
}
