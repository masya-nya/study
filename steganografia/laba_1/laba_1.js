let {Document} = require('docxyz');
let fileName = 'variant04.docx';
let document = new Document(fileName);

const RGBCheck = ({ r, g, b }) => {
    return Boolean(r && g && b)
}


let kod = '';
let shifrText = [];
for(let paragraph of document.paragraphs){
    for(let run of paragraph.runs) {
        const fontColor = RGBCheck(run.font.color.rgb);
        const fontSize = run.font.size.pt
        const fontHighlightColor = run.font.highlight_color
        const fontScale = run._r.get_or_add_rPr().xpath("./w:w")[0]
        const fontSpacing = run._r.get_or_add_rPr().xpath("./w:spacing")[0]
        if(
            fontColor ||
            (fontSize !== 12) ||
            (fontHighlightColor !== 8) ||
            fontScale ||    
            fontSpacing
            ) {
            shifrText.push(run.text);
            run.underline = true;
            for(let i = 0; i < run.text.length; i++) {
                kod += '1'
            }
        }
        else {
            for(let i = 0; i < run.text.length; i++) {
                kod += '0'
            }
        }
    }
}
// kod += '000'
console.log('<------------------------------------------------------------->')
console.log(kod)
console.log(kod.length)
console.log('<------------------------------------------------------------->')
console.log(shifrText)
const byteCharacters = kod.match(/.{1,8}/g)
// console.log('<------------------------------------------------------------->')
// console.log(byteCharacters)
const byteNumbers = byteCharacters.map(byte => parseInt(byte, 2))
// console.log('<------------------------------------------------------------->')
// console.log(byteNumbers)
const byteArray = new Uint8Array(byteNumbers)
const decoderWindows = new TextDecoder('windows-1251')
const decoderKOI8 = new TextDecoder('KOI8-R')
const decodercp866 = new TextDecoder('cp866')
console.log('KOI8-R --- |', decoderKOI8.decode(byteArray), '|')
console.log('<------------------------------------------------------------->')
console.log('cp866 --- |', decodercp866.decode(byteArray), '|')
console.log('<------------------------------------------------------------->')
console.log('Windows-1251 --- |', decoderWindows.decode(byteArray), '|')
console.log('<------------------------------------------------------------->')

document.save(fileName)