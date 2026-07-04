// तुमची चालू असलेली (सध्याची) Google Apps Script ची URL
const SCRIPT_URL = "https://script.google.com/macros/s/AKfycby0jkH1Cbmhq1QEFC04ekzTazjMJlG5FPQ9UPa4QRQj_vlrX26bBqm3Zvze8NrHFu5k/exec"; 

let parsedRetailerRows = null;
let parsedWholesalerRows = null;

// HTML पूर्ण लोड झाल्यावरच हे फंक्शन चालेल (Safe Mode)
document.addEventListener('DOMContentLoaded', function() {

    function processExcelFile(file, type, statusElem) {
        const reader = new FileReader();
        reader.onload = function(evt) {
            try {
                let workbook;
                if (file.name.toLowerCase().endsWith('.csv')) {
                    const textData = new TextDecoder("utf-8").decode(evt.target.result);
                    workbook = XLSX.read(textData, { type: 'string' });
                } else {
                    const data = new Uint8Array(evt.target.result);
                    workbook = XLSX.read(data, { type: 'array' });
                }

                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
                
                let firstRowText = (rawData[0] || []).join(" ").toLowerCase();
                
                if (type === 'retailer' && !firstRowText.includes("retailer")) {
                    statusElem.textContent = "❌ ही रिटेलरची फाईल नाही!";
                    statusElem.className = "file-status status-error";
                    parsedRetailerRows = null;
                    checkBothFiles();
                    return;
                }
                if (type === 'wholesaler' && !firstRowText.includes("wholesaler")) {
                    statusElem.textContent = "❌ ही होलसेलरची फाईल नाही!";
                    statusElem.className = "file-status status-error";
                    parsedWholesalerRows = null;
                    checkBothFiles();
                    return;
                }

                let headerRowIndex = -1;
                let searchKeyword = type === 'retailer' ? "district" : "district name";
                
                for(let i = 0; i < rawData.length; i++){
                    if(!rawData[i]) continue;
                    if(rawData[i].join(" ").toLowerCase().includes(searchKeyword)) {
                        headerRowIndex = i;
                        break;
                    }
                }

                if(headerRowIndex === -1) {
                    statusElem.textContent = "❌ योग्य कॉलम सापडला नाही!";
                    statusElem.className = "file-status status-error";
                    if(type==='retailer') parsedRetailerRows = null; else parsedWholesalerRows = null;
                    checkBothFiles();
                    return;
                }

                const headers = rawData[headerRowIndex].map(h => String(h || "").trim());
                const jsonRows = [];
                
                for(let i = headerRowIndex + 1; i < rawData.length; i++){
                    let row = rawData[i];
                    if(row.length === 0 || !row.join("").trim()) continue; 
                    
                    let obj = {};
                    for(let j = 0; j < headers.length; j++){
                        if(headers[j]) obj[headers[j]] = row[j] !== undefined ? row[j] : "";
                    }
                    let distValue = obj['District'] || obj['District Name'];
                    if(distValue && String(distValue).trim() !== "") {
                        jsonRows.push(obj);
                    }
                }

                if(jsonRows.length === 0) {
                    statusElem.textContent = "❌ फाईलमध्ये डेटा नाही.";
                    statusElem.className = "file-status status-error";
                    if(type==='retailer') parsedRetailerRows = null; else parsedWholesalerRows = null;
                } else {
                    statusElem.textContent = `✅ तयार! (${jsonRows.length} रेकॉर्ड्स)`;
                    statusElem.className = "file-status status-ok";
                    if(type==='retailer') parsedRetailerRows = jsonRows; else parsedWholesalerRows = jsonRows;
                }
                
                checkBothFiles();

            } catch (error) {
                statusElem.textContent = "❌ फाईल वाचण्यात त्रुटी!";
                statusElem.className = "file-status status-error";
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function checkBothFiles() {
        const uploadBtn = document.getElementById('uploadBtn');
        const msgWin = document.getElementById('msgWindow');
        if (parsedRetailerRows && parsedWholesalerRows) {
            uploadBtn.disabled = false;
            uploadBtn.innerHTML = "दोन्ही फाईल्स अपलोड करा";
            msgWin.textContent = "दोन्ही फाईल्स यशस्वीरित्या वाचल्या. आता डेटा अपलोड करा.";
            msgWin.className = "message success-msg";
        } else {
            uploadBtn.disabled = true;
            uploadBtn.innerHTML = "दोन्ही फाईल्स अपलोड करा";
            msgWin.className = "message";
        }
    }

    document.getElementById('retailerFile').addEventListener('change', function(e) {
        if(e.target.files[0]) processExcelFile(e.target.files[0], 'retailer', document.getElementById('retailerStatus'));
    });

    document.getElementById('wholesalerFile').addEventListener('change', function(e) {
        if(e.target.files[0]) processExcelFile(e.target.files[0], 'wholesaler', document.getElementById('wholesalerStatus'));
    });

    document.getElementById('uploadBtn').addEventListener('click', function() {
        if (!parsedRetailerRows || !parsedWholesalerRows) return;

        const msgWin = document.getElementById('msgWindow');
        const uploadBtn = document.getElementById('uploadBtn');

        msgWin.className = "message success-msg";
        msgWin.textContent = "प्रक्रिया सुरू आहे... डेटा अपडेट होत आहे, कृपया थांबा.";
        
        // Loading animation
        uploadBtn.disabled = true;
        uploadBtn.innerHTML = '<span class="loader"></span> अपलोड होत आहे...';

        const payload = {
            action: 'uploadAllStock',
            retailerRows: parsedRetailerRows,
            wholesalerRows: parsedWholesalerRows
        };

        fetch(SCRIPT_URL, {
            method: 'POST',
            body: JSON.stringify(payload)
        })
        .then(response => response.json()) 
        .then(data => {
            uploadBtn.innerHTML = "दोन्ही फाईल्स अपलोड करा";
            if (data.success) {
                msgWin.className = "message success-msg";
                msgWin.textContent = "✅ " + data.message;
                
                parsedRetailerRows = null;
                parsedWholesalerRows = null;
                document.getElementById('retailerFile').value = "";
                document.getElementById('wholesalerFile').value = "";
                document.getElementById('retailerStatus').textContent = "फाईल निवडलेली नाही";
                document.getElementById('retailerStatus').className = "file-status";
                document.getElementById('wholesalerStatus').textContent = "फाईल निवडलेली नाही";
                document.getElementById('wholesalerStatus').className = "file-status";
            } else {
                msgWin.className = "message error-msg";
                msgWin.textContent = "❌ त्रुटी: " + data.message;
                uploadBtn.disabled = false;
            }
        })
        .catch(error => {
            uploadBtn.innerHTML = "दोन्ही फाईल्स अपलोड करा";
            uploadBtn.disabled = false;
            msgWin.className = "message error-msg";
            msgWin.textContent = "❌ नेटवर्क त्रुटी. सर्व्हरशी संपर्क होऊ शकला नाही.";
        });
    });
});
