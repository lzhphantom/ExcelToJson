# ExcelToJson
exchange excel info to json

```
result,err := AllSheetToJson(filePath)
```

AllSheetToJson

```json
{
        "Sheet1": [
            {
                "age": "20",
                "name": "zs",
                "sex": "man"
            },
            {
                "age": "19",
                "name": "ls",
                "sex": "man"
            },
            {
                "age": "17",
                "name": "emz",
                "sex": "woman"
            }
        ],
        "Sheet2": [
            {
                "age": "20",
                "name": "zs",
                "tel": "18328846666"
            },
            {
                "age": "19",
                "name": "ls",
                "tel": "18328846667"
            },
            {
                "age": "17",
                "name": "emz",
                "tel": "18328846668"
            }
        ]
    }
```

```
ASheetToJson(filePath, sheetName)
```

``` json
[
        {
            "age": "20",
            "name": "zs",
            "sex": "man"
        },
        {
            "age": "19",
            "name": "ls",
            "sex": "man"
        },
        {
            "age": "17",
            "name": "emz",
            "sex": "woman"
        }
]
```