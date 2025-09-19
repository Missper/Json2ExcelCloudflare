# Json转多Sheet的Excel表格

## 示例输入

`0-9`和`A`会各自创建一个sheet, 其中的`Product Name`等会作为表头

```json
{
  "0-9": [
    {
      "Product Name": "0.015% CHLORHEXIDINE & 0.15% CETRIMIDE SOL",
      "Permit No.": "HK42682",
      "Active Ingredients": "CETRIMIDE,CHLORHEXIDINE ACETATE",
      "Company Name": "BAXTER HEALTHCARE LTD",
      "Company Address": "RM 2701-2, OXFORD HOUSE TAIKOO PLACE, 979 KING'S ROAD ISLAND EAST, HONG KONG"
    },
    {
      "Product Name": "0.05% CHLORHEXIDINE & 0.5% CETRIMIDE SOL",
      "Permit No.": "HK42684",
      "Active Ingredients": "CETRIMIDE,CHLORHEXIDINE ACETATE",
      "Company Name": "BAXTER HEALTHCARE LTD",
      "Company Address": "RM 2701-2, OXFORD HOUSE TAIKOO PLACE, 979 KING'S ROAD ISLAND EAST, HONG KONG"
    }
  ],
  "A": [
    {
      "Product Name": "A P CAP 250MG",
      "Permit No.": "HK10610",
      "Active Ingredients": "AMPICILLIN (AS TRIHYDRATE)",
      "Company Name": "THE UNITED LABORATORIES LTD",
      "Company Address": "YUEN LONG INDUSTRIAL ESTATE, 6 FUK WANG ST YUEN LONG NT"
    },
    {
      "Product Name": "A SCABS LOTION 5%",
      "Permit No.": "HK56053",
      "Active Ingredients": "PERMETHRIN",
      "Company Name": "TAISHO PHARMACEUTICAL (HK) LTD",
      "Company Address": "RM A, 7 FLOOR, HARVEST MOON HOUSE, 337-339 NATHAN ROAD, JORDAN, KOWLOON"
    }
  ]
}

```



## 输出

二进制文件, excel格式, 文件名data.xlsx





## 部署方式

```bash
# 💻 Continue Developing
# Change directories: 
cd testworker

# Start dev server: 
npm run start

# Deploy: 
npm run deploy
```



