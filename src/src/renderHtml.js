var __defProp = Object.defineProperty;
var __name = (target, value) => __defProp(target, "name", { value, configurable: true });

// src/templates/populated-worker/src/renderHtml.js
function renderHtml(content) {
  return `
    <!DOCTYPE html>
    <html lang="en">
      <head>
        <meta charset="UTF-8" />
        <meta name="viewport" content="width=device-width, initial-scale=1.0" />
        <title>D1</title>
        <link rel="stylesheet" type="text/css" href="https://templates.cloudflareintegrations.com/styles.css">
      </head>
    
      <body>
        <header>
          <img
            src="https://imagedelivery.net/wSMYJvS3Xw-n339CbDyDIA/30e0d3f6-6076-40f8-7abb-8a7676f83c00/public"
          />
          <h1>Successfully connected testworker</h1>
        </header>
        <main>
          <p>This worker is set up to handle POST requests with JSON bodies.</p>
        </main>  
      </body>
    </html>
`;
}


__name(renderHtml, "renderHtml");
export {
  renderHtml as default
};
