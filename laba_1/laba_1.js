let {Document} = require('docxyz');
let fileName = 'variant04.docx';
let document = new Document(fileName);
let a = [];

const RGBCheck = ({ r, g, b }) => {
    return Boolean(r && g && b)
}

let kod = '';

for(let paragraph of document.paragraphs){
    for(let run of paragraph.runs) {
        const fontColor = RGBCheck(run.font.color.rgb);
        const fontSize = run.font.size.pt
        const fontHighlightColor = run.font.highlight_color
        const fontScale = run._r.get_or_add_rPr().xpath("./w:w")[0]
        const fontSpacing = run._r.get_or_add_rPr().xpath("./w:spacing")[0]
        // console.log(fontSpacing[0])
        // console.log(run.text)
        if(
            fontColor ||
            (fontSize !== 12) ||
            (fontHighlightColor !== 8) ||
            fontScale ||
            fontSpacing
            ) {

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

console.log(kod)
console.log(kod.length)
const byteCharacters = kod.match(/.{1,8}/g)
console.log(byteCharacters)
const byteNumbers = byteCharacters.map(byte => parseInt(byte, 2))
console.log(byteNumbers)
const byteArray = new Uint8Array(byteNumbers)
const decoderWindows = new TextDecoder('windows-1251')
const decoderKOI8 = new TextDecoder('KOI8-R')
const decodercp866 = new TextDecoder('cp866')
console.log('KOI8-R --- |', decoderKOI8.decode(byteArray), '|')
console.log('cp866 --- |', decodercp866.decode(byteArray), '|')
console.log('Windows-1251 --- |', decoderWindows.decode(byteArray), '|')

