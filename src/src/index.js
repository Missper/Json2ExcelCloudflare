// src/templates/populated-worker/src/index.js
import renderHtml from "./renderHtml.js";
import ExcelJS from 'exceljs'; // 确保这个库能被您的 Worker 访问到

async function createExcelFromJSON(jsonData) {

    console.log("开始处理数据...", typeof jsonData);
    console.log("收到的数据:", JSON.stringify(jsonData).substring(0, 1000));
    const workbook = new ExcelJS.Workbook();
    
    // 假设jsonData的第一个元素就是要处理的对象
    const sheetData = jsonData;
    console.log("解析后的 sheetData 类型:", typeof sheetData);
    console.log("解析后的 sheetData:", JSON.stringify(sheetData).substring(0, 1000));

    // 遍历 sheetData 对象中的所有键
    for (const sheetName in sheetData) {
        if (Object.prototype.hasOwnProperty.call(sheetData, sheetName)) {
            const dataRows = sheetData[sheetName];
            console.log(`正在处理工作表: ${sheetName}`);
            console.log(`数据行数: ${Array.isArray(dataRows) ? dataRows.length : 0}`);

            // 检查数据是否为非空数组
            if (Array.isArray(dataRows) && dataRows.length > 0) {
                // 创建一个新的工作表，名称为键名
                const worksheet = workbook.addWorksheet(sheetName);
                console.log(`已创建工作表: ${sheetName}`);

                // 从第一行数据中提取列名作为表头
                const columns = Object.keys(dataRows[0]).map(key => ({
                    header: key,
                    key: key,
                    width: key.length + 5 // 自动设置列宽，确保可见性
                }));
                worksheet.columns = columns;
                console.log(`设置表头: ${JSON.stringify(columns)}`);

                // 填充数据行
                worksheet.addRows(dataRows);
                console.log(`已添加数据行`);
            } else {
                console.log(`工作表 ${sheetName} 没有有效数据，跳过`);
            }
        }
    }

    // 将工作簿写入缓冲区，返回一个 Blob 或 Buffer
    const buffer = await workbook.xlsx.writeBuffer();
    console.log("Excel 文件已生成，准备返回响应");

    return new Response(buffer, {
        headers: {
            'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            'Content-Disposition': 'attachment; filename="data.xlsx"'
        }
    });
}

var src_default = {
  async fetch(request, env) {
    console.log("收到请求:", request.method);
    if (request.method === "POST") {
      try {
        const requestBody = await request.json();
        // console.log("POST 请求体:", JSON.stringify(requestBody));
        
        // 假设 requestBody 的格式与您的 JSON 结构匹配
        return await createExcelFromJSON(requestBody);

      } catch (error) {
        // 如果解析 JSON 或生成 Excel 失败，返回一个错误
        console.error("处理 POST 请求时出错:", error);
        return new Response(JSON.stringify({ error: error.message || "Invalid JSON or excel generation failed" }), {
          status: 400,
          headers: { "content-type": "application/json" }
        });
      }
    }

    // 处理 GET 请求
    console.log("处理 GET 请求，返回 HTML 页面");
    return new Response(
      renderHtml("<h1>欢迎！请发送一个 POST 请求来生成 Excel。</h1>"),
      {
        headers: { "content-type": "text/html" }
      }
    );
  }
};

export { src_default as default };