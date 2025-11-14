const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = 3000;

// Tạo thư mục uploads nếu chưa có
if (!fs.existsSync('./uploads')) fs.mkdirSync('./uploads');

// Cấu hình multer
const storage = multer.diskStorage({
    destination: (req, file, cb) => cb(null, 'uploads/'),
    filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname)
});
const upload = multer({ storage });

app.use(express.static('public'));

// Thứ tự công đoạn (quan trọng để xác định công đoạn cuối)
const processesOrder = [
    "Extrusion",
    "UV",
    "UV+Bigsheet",
    "Profiling",
    "Profiling+Bevel",
    "Packaging",
    "Padding+Packaging"
];

// Mảng tên chi phí
const chiPhiNames = [
    { no: 1, code: '01', name: 'Chi phí nhân công trực tiếp' },
    { no: 2, code: '02', name: 'Chi phí nhân công gián tiếp' },
    { no: 3, code: '03', name: 'Chi phí dụng cụ sản xuất' },
    { no: 4, code: '04', name: 'Chi phí khấu hao TSCĐ' },
    { no: 5, code: '05', name: 'Chi phí dịch vụ mua ngoài' },
    { no: 6, code: '06', name: 'Chi phí bằng tiền khác' },
];

// Hàm parse chi phí
function parseChiPhi(value) {
    if (!value) return 0;
    return parseFloat(value.toString().replace(/^'+/,'').replace(/,/g,'').trim()) || 0;
}

// Đọc chi phí
function readChiPhiCOCT(sheet, rowNumber) { return ['CO','CP','CQ','CR','CS','CT'].map(c => parseChiPhi(sheet[c+rowNumber]?.v)); }
function readChiPhiUv(sheet, rowNumber) { return ['CU','CV','CW','CX','CY','CZ'].map(c => parseChiPhi(sheet[c+rowNumber]?.v)); }
function readChiPhiProfiling(sheet, rowNumber) { return ['DA','DB','DC','DD','DE','DF'].map(c => parseChiPhi(sheet[c+rowNumber]?.v)); }
function readChiPhiPackaging(sheet, rowNumber) { return ['DY','DZ','EA','EB','EC','ED'].map(c => parseChiPhi(sheet[c+rowNumber]?.v)); }
function readChiPhiUVBigsheet(sheet,rowNumber){ 
    const colsPairs = [['CU','DS'],['CV','DT'],['CW','DU'],['CX','DV'],['CY','DW'],['CZ','DX']];
    return colsPairs.map(([c1,c2]) => parseChiPhi(sheet[c1+rowNumber]?.v)+parseChiPhi(sheet[c2+rowNumber]?.v));
}
function readChiPhiPaddingPackaging(sheet,rowNumber){
    const colsPairs = [['DM','DY'],['DN','DZ'],['DO','EA'],['DP','EB'],['DQ','EC'],['DR','ED']];
    return colsPairs.map(([c1,c2]) => parseChiPhi(sheet[c1+rowNumber]?.v)+parseChiPhi(sheet[c2+rowNumber]?.v));
}
function readChiPhiProfilingBevel(sheet,rowNumber){
    const colsPairs = [['DA','DG'],['DB','DH'],['DC','DI'],['DD','DJ'],['DE','DK'],['DF','DL']];
    return colsPairs.map(([c1,c2]) => parseChiPhi(sheet[c1+rowNumber]?.v)+parseChiPhi(sheet[c2+rowNumber]?.v));
}

app.post('/upload', upload.fields([{ name: 'routing' }, { name: 'resource' }]), (req, res) => {
    let routingFile, resourceFile, outFile;
    try {
        routingFile = req.files['routing'][0].path;
        resourceFile = req.files['resource'][0].path;

        const resourceWB = xlsx.readFile(resourceFile);
        const routingWB = xlsx.readFile(routingFile);

        const resourceSheet = resourceWB.Sheets['CP CONG DOAN THEO DATA'];
        const routingSheet = routingWB.Sheets[routingWB.SheetNames[0]];

        // --- Resource ---
        const resourceRange = xlsx.utils.decode_range(resourceSheet['!ref']);
        const resourceData = [];
        for (let r=4; r<=resourceRange.e.r+1; r++){
            const maTinAn = resourceSheet['C'+r]?.v?.toString().trim() || '';
            const version = resourceSheet['D'+r]?.v?.toString().replace(/^'+/,'').trim() || '';
            resourceData.push({ maTinAn, version, rowNumber: r });
        }

        // --- Routing ---
        const routingRange = xlsx.utils.decode_range(routingSheet['!ref']);
        const routingData = [];
        for(let r=2;r<=routingRange.e.r+1;r++){
            const routeKeyA = routingSheet['A'+r]?.v?.toString().trim() || '';
            const inventoryID = routingSheet['B'+r]?.v?.toString().trim() || '';
            const routingNo = routingSheet['G'+r]?.v?.toString().trim() || '';
            const routeVersion = routingSheet['D'+r]?.v?.toString().replace(/^'+/,'').trim() || '';
            const congDoan = routingSheet['H'+r]?.v?.toString().trim() || '';
            routingData.push({routeKeyA, inventoryID, routingNo, routeVersion, congDoan});
        }

        // --- Tìm công đoạn cuối theo từng mã đầu 5 ---
        const lastProcessMap = {};   
        routingData.forEach(r => {
            const key = r.routeKeyA;
            if (!key) return;
            const idx = processesOrder.indexOf((r.congDoan || "").trim());
            if (idx === -1) return;
            if (!lastProcessMap[key] || idx > processesOrder.indexOf(lastProcessMap[key])) {
                lastProcessMap[key] = r.congDoan;
            }
        });

        // --- Gán chi phí ---
        const updatedData = routingData.map(route => {
            let chiPhiArr=[0,0,0,0,0,0];
            const resRows = resourceData.filter(r=>r.maTinAn===route.routeKeyA);
            if(resRows.length>0){
                let resRow = resRows.find(r=>r.version===route.routeVersion);
                if(!resRow){ 
                    resRows.sort((a,b)=>a.version.localeCompare(b.version));
                    resRow = resRows[0];
                }
                const lowerCD = (route.congDoan||'').toLowerCase();
                if(lowerCD==='extrusion') chiPhiArr=readChiPhiCOCT(resourceSheet,resRow.rowNumber);
                else if(lowerCD==='uv+bigsheet') chiPhiArr=readChiPhiUVBigsheet(resourceSheet,resRow.rowNumber);
                else if(lowerCD==='profiling') chiPhiArr=readChiPhiProfiling(resourceSheet,resRow.rowNumber);
                else if(lowerCD==='packaging') chiPhiArr=readChiPhiPackaging(resourceSheet,resRow.rowNumber);
                else if(lowerCD==='uv') chiPhiArr=readChiPhiUv(resourceSheet,resRow.rowNumber);
                else if(lowerCD==='padding+packaging') chiPhiArr=readChiPhiPaddingPackaging(resourceSheet,resRow.rowNumber);
                else if(lowerCD==='profiling+bevel') chiPhiArr=readChiPhiProfilingBevel(resourceSheet,resRow.rowNumber);
            }
            return {...route, chiPhiArr};
        });

        // --- Long Format ---
        const longData=[];
        updatedData.forEach(route=>{

            const lastCD = (lastProcessMap[route.routeKeyA] || "").toLowerCase();
            const isLast = lastCD === (route.congDoan || "").toLowerCase();
            const notPackaging = !["packaging","padding+packaging"].includes(lastCD);

            // APPLY LOGIC GIỐNG HỆ ROUTING CŨ
            let finalInventoryID = route.inventoryID;
            if (isLast && notPackaging) {
                finalInventoryID = route.routeKeyA;
            }

            chiPhiNames.forEach((cp,i)=>{
                let value='';
                const lowerCD=(route.congDoan||'').toLowerCase();
                if(['extrusion','uv+bigsheet','profiling','packaging','uv','padding+packaging','profiling+bevel'].includes(lowerCD))
                    value=route.chiPhiArr[i];

                longData.push({
                    'Mã Đầu 5': route.routeKeyA,
                    'InventoryID': finalInventoryID,
                    'Routing Version': route.routeVersion,
                    'Routing No': route.routingNo,
                    'Routing Name': route.congDoan,
                    'No': cp.no,
                    'Resource CD': cp.code,
                    'Resource Name': cp.name,
                    'Price': value
                });
            });
        });

        // --- Xuất Excel ---
        const newWB = xlsx.utils.book_new();
        const newWS = xlsx.utils.json_to_sheet(longData);
        xlsx.utils.book_append_sheet(newWB,newWS,'LongFormat');
        outFile = path.join(__dirname,'uploads','Resource_import.xlsx');
        xlsx.writeFile(newWB,outFile);

        res.download(outFile,err=>{
            if(err) console.error("Lỗi khi gửi file:",err);
            [routingFile,resourceFile,outFile].forEach(f=>{if(f&&fs.existsSync(f)) try{fs.unlinkSync(f);}catch(e){}});
        });

    }catch(err){
        console.error(err);
        [routingFile,resourceFile,outFile].forEach(f=>{if(f&&fs.existsSync(f)) try{fs.unlinkSync(f);}catch(e){}});
        res.status(500).send('Có lỗi xảy ra khi xử lý file: '+err.message);
    }
});

app.listen(PORT,()=>console.log(`Server running at http://localhost:${PORT}`));
