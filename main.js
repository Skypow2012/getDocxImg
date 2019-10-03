const unzip = require('unzip');
const fs = require('fs');

const argv = process.argv.splice(2);
if (!argv[0]) {
    console.log('请输入待转换目录 或 文件名！');
}
let inputStr = argv.join(' ');
if (inputStr.indexOf('.docx') > -1) {
    const needGetFileName = inputStr.replace(/\\/g,'\/');
    getDocxImage(needGetFileName);
} else {
    const needGetDirPath = inputStr.replace(/\\/g,'\/');
    getDirImages(needGetDirPath)
}
function getDirImages(dirPath) {
    let fileNames = fs.readdirSync(dirPath);
    for (let i = 0; i < fileNames.length; i++) {
        const fileName = fileNames[i];
        if (fileName.indexOf('.docx')>-1) {
            getDocxImage(`${dirPath}/${fileName}`);
        }
    }
}

function getText(xml) {
    xml = xml.replace(/<\/w:p>/g, '\r\n');
    let i = 0;
    xml = xml.replace(/<[^<>]*?docPr[^<>]*?id="(\d+)"[^<>]*?>/g, 'Image_$1');
    xml = xml.replace(/<w:b\/>/g, '【加粗】');
    xml = xml.replace(/<[^<>]*?center[^<>]*?>/g, '【居中】');
    let text = xml.replace(/<\/*.*?>/g, '');
    text = text.replace(/([^>]{1})【加粗】/g, '$1');
    text = text.replace(/\r\n */g, '\r\n');

    text = text.replace(/供图/g, '摄');
    text = text.replace(/\r\n([^【]*)\r\n/g, '\r\n【普通】$1\r\n');
    return text;
}
function fixText2Xml(text) {
    text = text.replace(/【普通】(.*?)\r\n/g, '<p style="text-indent: 32px;"><span style="font-family:宋体">$1</span></p>\r\n')
    text = text.replace(/【居中】【加粗】(.*?)\r\n/g, '<p style="text-align:center"><span style="font-family:黑体">$1</span></p>\r\n')
    text = text.replace(/【居中】(.*?)\r\n/g, '<p style="text-align:center"><span style="font-family:宋体">$1</span></p>\r\n')
    text = text.replace(/【加粗】(.*?)\r\n/g, '<p><span style="font-family:黑体">$1</span></p>\r\n')
    text = text.replace(/Image_(\d*)/g, `
    <img src="./demo/word/media/image$1.jpeg") no-repeat center center;border:1px solid #ddd"/>
    `);//width="285" height="426" 
    // fs.writeFileSync('1.html', text)
    return text;
}
function getDocxImage(imagePath) {
    const parentDirPath = imagePath.match(/(.*)[\/\\].*/)[1];;
    const fileName = imagePath.match(/.*[\/\\](.*)/)[1];// '植物文章样例.docx';
    
    const tempDirName = 'temp' + new Date().getTime();
    const tarDirName = fileName.match(/(.*)\..*/)[1];
    
    const f = fs.createReadStream(`${parentDirPath}/${fileName}`)
    f.pipe(unzip.Extract({ path: `${parentDirPath}/${tempDirName}` }))
    function deleteFolder(path) {
        var files = [];
        if( fs.existsSync(path) ) {
            files = fs.readdirSync(path);
            files.forEach(function(file,index){
                var curPath = path + "/" + file;
                if(fs.statSync(curPath).isDirectory()) { // recurse
                    deleteFolder(curPath);
                } else { // delete file
                    fs.unlinkSync(curPath);
                }
            });
            fs.rmdirSync(path);
        }
    };
    let inter = setInterval(() => {
        if (!fs.existsSync(`${parentDirPath}/${tempDirName}`)) {
            return;
        } else {
            clearInterval(inter);
        }
        let images = fs.readdirSync(`${parentDirPath}/${tempDirName}/word/media`);
        let wordXml = fs.readFileSync(`${parentDirPath}/${tempDirName}/word/document.xml`).toString();
        // console.log(wordXml.toString())
        // let expReg = /<w:t>([^<>]*?摄[^<>]*?)<\/w:t>/g;
        let xmlText = getText(wordXml);
        let expReg = /Image_\d*\r\n(.*?)\r\n/g;
        let matched = '';
        let matches = [];
        while (matched = expReg.exec(xmlText)) {
            matches.push(matched[1].replace(/【.*?】/g, ''));
        }
        console.log(matches);
        if (!fs.existsSync(`${parentDirPath}/${tarDirName}`)) {
            fs.mkdirSync(`${parentDirPath}/${tarDirName}`);
        }
        for (let i = 0; i < images.length; i++) {
            const imageName = images[i];
            const imageExtName = imageName.match(/.*\.(.*)/)[1];
            let imageBin = fs.readFileSync(`${parentDirPath}/${tempDirName}/word/media/${imageName}`);
            let tarName = matches[i];
            fs.writeFileSync(`${parentDirPath}/${tarDirName}/${tarName}.${imageExtName}`, imageBin);
            console.log(` 生成 ${parentDirPath}/${tarDirName}/${tarName}.${imageExtName}`);
        }
        // deleteFolder(`${parentDirPath}/${tempDirName}`);
        console.log(`成功获取 ${fileName}`);
    }, 1000);
}