# ExcelOverwriter

Overwrite the data in one XLSX using data from another XLSX, using column headers that you specify.

Initialize it like this:

    new XlsxSetup()
    new UpdateByHeaders(
    /*Input Path*/		"input.xlsx",
    /*Input Sheet Regex*/	~/thisIsMySheet/,
    /*Input Key Regex*/	~/thisIsMyKey/,
    /*Output Path*/		"output.xlsx",
    /*Output Sheet Regex*/	~/thisIsMyOtherSheet/,
    /*Output Key Regex*/	~/thisIsMyOtherKey/,
    /*In-to-Out Headers*/	[
    "thisIsMyKey": "thisIsMyOtherKey",
    "Apples and Oranges": "Fruit",
    "Veggies": "Carrots and Tomatoes"
    ]
    )
