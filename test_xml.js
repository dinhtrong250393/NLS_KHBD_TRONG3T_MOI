const docXml = "<w:p>hello</w:p><w:p>world</w:p><w:p>test</w:p>";
let pArray = docXml.split("</w:p>");
let pTexts = pArray.map(p => p.replace(/<[^>]+>/g, ''));
console.log(pTexts);
