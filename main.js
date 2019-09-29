const unzip = require('unzip');
const fs = require('fs');

const argv = process.argv.splice(2);
if (!argv[0]) {
    console.log('请输入待转换目录 或 文件名！');
}
let inputStr = argv.join(' ');
if (inputStr.indexOf('.docx') > -1) {
    const needGetFileName = argv[0];
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
        let expReg = /<w:t>([^<>]*?摄[^<>]*?)<\/w:t>/g;
        let matched = '';
        let matches = [];
        while (matched = expReg.exec(wordXml)) {
            matches.push(matched[1]);
        }
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
        deleteFolder(`${parentDirPath}/${tempDirName}`);
        console.log(`成功获取 ${fileName}`);
    }, 1000);
}