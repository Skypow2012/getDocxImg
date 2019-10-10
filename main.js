const unzip = require('unzip');
const fs = require('fs');

const argv = process.argv.splice(2);
if (!argv[0]) {
    console.log('请输入待转换目录 或 文件名！');
}
let inputStr = argv.join(' ');
if (inputStr.indexOf('.docx') > -1) {
    const needGetFileName = inputStr.replace(/\\/g, '\/');
    getDocxImage(needGetFileName);
} else if (inputStr.indexOf('.doc') > -1) {
    const needGetFileName = inputStr.replace(/\\/g, '\/');
    tranDoc2Docx(needGetFileName, getDocxImage);
} else {
    const needGetDirPath = inputStr.replace(/\\/g, '\/');
    getDirImages(needGetDirPath)
}
function getDirImages(dirPath) {
    let fileNames = fs.readdirSync(dirPath);
    for (let i = 0; i < fileNames.length; i++) {
        const fileName = fileNames[i];
        if (fileName.indexOf('.bak.docx') > -1) {
            console.log('delete', fileName);
            fs.unlinkSync(`${dirPath}/${fileName}`)
        } else if (fileName.indexOf('.docx') > -1) {
            console.log(fileName, 'fileName')
            getDocxImage(`${dirPath}/${fileName}`);
        } else if (fileName.indexOf('.doc') > -1) {
            console.log(fileName, 'fileName')
            tranDoc2Docx(`${dirPath}/${fileName}`, getDocxImage);
        } else if (fileName.indexOf('temp') == 0 && fs.statSync(`${dirPath}/${fileName}`).isDirectory()) {
            deleteFolder(`${dirPath}/${fileName}`);
        }
    }
}

function tranDoc2Docx(path, cb) {
    var exec = require('child_process').exec;
    function execute() {
        var cmd = `py ${__dirname}/py/demo.py ${path}`;
        console.log(cmd);
        exec(cmd, function (error, stdout, stderr) {
            if (error) {
                console.log(error);
            } else {
                console.log("成功转换", path);
                let tarPath = path.replace('.doc', '.bak.docx')
                cb(tarPath);
            }
        });
    }
    execute();
}

function getText(xml) {
    xml = xml.replace(/instrText.*?<\/w:instrText/g, '');
    xml = xml.replace(/<\/w:p>/g, '\r\n');
    let i = 0;
    // xml = xml.replace(/<[^<>]*?docPr[^<>]*?id="(\d+)"[^<>]*?>/g, 'Image_$1');
    // if (xml.indexOf('Image_' == -1)) {
    //     xml = xml.replace(/<[^<>]*?shape id="图片 (\d+)"[^<>]*?>/g, 'Image_$1');
    // }
    while (xml.indexOf('<wp:docPr ') > -1 || xml.indexOf('<v:shape ') > -1) {
        xml = xml.replace(/<(wp:docPr |v:shape )[^<>]*?>/, 'Image_' + ++i);
    }
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
    const fileName = imagePath.match(/.*[\/\\](.*)/)[1]; // '植物文章样例.docx';
    const tempDirName = 'temp' + new Date().getTime();
    let tarDirName = fileName.match(/(.*?)\..*/)[1];

    const f = fs.createReadStream(`${parentDirPath}/${fileName}`)
    f.pipe(unzip.Extract({ path: `${parentDirPath}/${tempDirName}` }))
    let inter = setInterval(() => {
        if (fs.existsSync(`${parentDirPath}/${tempDirName}`) && fs.existsSync(`${parentDirPath}/${tempDirName}/word/document.xml`)) {
            clearInterval(inter);
        } else {
            return;
        }
        let images = []
        let mediaPath = `${parentDirPath}/${tempDirName}/word/media`
        if (fs.existsSync(mediaPath)) {
            images = fs.readdirSync(mediaPath);
        }
        if (!images.length) {
            tarDirName = '无图-' + tarDirName;
        }
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
            let tarName = matches[i] || `image_${i + 1}`;
            fs.writeFileSync(`${parentDirPath}/${tarDirName}/${tarName}.${imageExtName}`, imageBin);
            console.log(` 生成 ${parentDirPath}/${tarDirName}/${tarName}.${imageExtName}`);
        }
        deleteFolder(`${parentDirPath}/${tempDirName}`);
        deleteBakFile(imagePath)
        console.log(`成功获取 ${fileName}`);
    }, 1000);
}

function deleteBakFile(filePath) {
    if (filePath.indexOf('.bak') > -1) {
        fs.unlinkSync(filePath);
    }
}

function deleteFolder(path) {
    var files = [];
    if (fs.existsSync(path)) {
        files = fs.readdirSync(path);
        files.forEach(function (file, index) {
            var curPath = path + "/" + file;
            if (fs.statSync(curPath).isDirectory()) { // recurse
                deleteFolder(curPath);
            } else { // delete file
                fs.unlinkSync(curPath);
            }
        });
        fs.rmdirSync(path);
    }
};