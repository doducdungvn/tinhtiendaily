'use strict';
(function() {
    let appMode = 'excel'; 
    let excelData1 = [], excelData2 = [];
    let rawDataMap = new Map(); 
    let rawDateRange = "Chưa có ngày";
    let dateRangeText = "Chưa có ngày";
    let regionDefinitions = {}, isRegionLocked = false; 
    let debtData = new Map(), isDebtFromExcel = false; 
    
    const REGION_DATA_URL = 'https://raw.githubusercontent.com/doducdungvn/tinhtiendaily/main/donvi.txt';
    
    const state = { currentSoSo: null, baseTongDT: 0, baseTongHH: 0, baseThuong: 0, extraBonus: 0, extraDebt: 0, addedTickets: [], accountCodes: new Set(), debtFromData: 0 };

    const DOM = {
        homePage: document.getElementById('homePage'), mainPage: document.getElementById('mainPage'),
        excelFile1: document.getElementById('excelFile1'), excelFile2: document.getElementById('excelFile2'), btnLoadData: document.getElementById('btnLoadData'), 
        
        tabExcel: document.getElementById('tabExcel'), tabRaw: document.getElementById('tabRaw'),
        panelExcel: document.getElementById('panelExcel'), panelRaw: document.getElementById('panelRaw'),
        rawInput: document.getElementById('rawInput'),
        
        btnGoToSummary: document.getElementById('btnGoToSummary'), btnGoToSell: document.getElementById('btnGoToSell'),
        btnBackFromSummary: document.getElementById('btnBackFromSummary'), btnBackFromSell: document.getElementById('btnBackFromSell'),

        summarySection: document.getElementById('summarySection'), reportSection: document.getElementById('reportSection'), sellSection: document.getElementById('sellSection'),
        summaryTableContainer: document.getElementById('summaryTableContainer'), summaryDate: document.getElementById('summaryDate'), summaryTitle: document.getElementById('summaryTitle'),
        soSoInput: document.getElementById('soSoInput'),
        
        // 2 Nút báo cáo
        btnCreateDailyReport: document.getElementById('btnCreateDailyReport'), btnCreateSummaryReport: document.getElementById('btnCreateSummaryReport'),
        
        reportContainer: document.getElementById('reportContainer'),
        actionButtonContainer: document.getElementById('actionButtonContainer'), actionChoices: document.getElementById('actionChoices'), toggleChoicesBtn: document.getElementById('toggleChoicesBtn'),
        btnCopySummary: document.getElementById('btnCopySummary'), btnPrintReport: document.getElementById('btnPrintReport'), btnExportExcel: document.getElementById('btnExportExcel'),
        btnCopyReport: document.getElementById('btnCopyReport'),
        btnAddTicket: document.getElementById('btnAddTicket'), sellQuantityInput: document.getElementById('sellQuantityInput'),
        manualInputModal: document.getElementById('manualInputModal'), btnShowManualReport: document.getElementById('btnShowManualReport'), btnShowManualReport_inline: document.getElementById('btnShowManualReport_inline'),
        btnCreateManualReport: document.getElementById('btnCreateManualReport'), btnClearManualReport: document.getElementById('btnClearManualReport'),
        manualAgentName: document.getElementById('manualAgentName'), manualDateRange: document.getElementById('manualDateRange'), manualLotoDB: document.getElementById('manualLotoDB'), manualLoCap: document.getElementById('manualLoCap'), manualC227: document.getElementById('manualC227'), manualC323: document.getElementById('manualC323'), manualThuong: document.getElementById('manualThuong'),
        sellTicketTypeRadios: document.querySelectorAll('input[name="sellTicketType"]'), sellPriceRadios: document.querySelectorAll('input[name="sellPrice"]'),
        customTicketName: document.getElementById('customTicketName'), customTicketPrice: document.getElementById('customTicketPrice'), customTicketRate: document.getElementById('customTicketRate'),
        sellRateRadios: document.querySelectorAll('input[name="sellRate"]'),
        
        closeModalBtn: document.getElementById('closeModalBtn'), notificationArea: document.getElementById('notification-area'), loadingIndicator: document.getElementById('loading-indicator'),
        imagePreviewModal: document.getElementById('imagePreviewModal'), closePreviewModalBtn: document.getElementById('closePreviewModalBtn'), imagePreviewHolder: document.getElementById('imagePreviewHolder'),
        logoContainers: document.querySelectorAll('.simple-animated-logo'), scrollToTopBtn: document.getElementById('scrollToTopBtn'), scrollToBottomBtn: document.getElementById('scrollToBottomBtn'),
        regionLockContainer: document.getElementById('regionLockContainer'), regionDefinitionsInput: document.getElementById('regionDefinitionsInput'), btnSaveRegions: document.getElementById('btnSaveRegions'), regionSelect: document.getElementById('regionSelect'),
        sortSelect: document.getElementById('sortSelect'), minRevenueInput: document.getElementById('minRevenueInput'), maxRevenueInput: document.getElementById('maxRevenueInput'), btnClearRevenueFilter: document.getElementById('btnClearRevenueFilter'),
        debtInput: document.getElementById('debtInput'), debtManualContainer: document.getElementById('debtManualContainer'), debtExcelContainer: document.getElementById('debtExcelContainer'), debtFile: document.getElementById('debtFile'), 
        debtModeRadios: document.querySelectorAll('input[name="debtMode"]'), btnClearDebt: document.getElementById('btnClearDebt'), debtStatusMsg: document.getElementById('debtStatusMsg'), btnApplyDebt: document.getElementById('btnApplyDebt'),
        passwordModal: document.getElementById('passwordModal'), closePasswordModalBtn: document.getElementById('closePasswordModalBtn'), passwordInput: document.getElementById('passwordInput'), btnSubmitPassword: document.getElementById('btnSubmitPassword'), cancelPasswordBtn: document.getElementById('cancelPasswordBtn')
    };
    
    const manualFormOrder = ['manualAgentName', 'manualDateRange', 'manualLotoDB', 'manualLoCap', 'manualC227', 'manualC323', 'manualThuong'];

    // --- KHỞI TẠO BAN ĐẦU: Ẩn nút Daily vì mặc định là Excel ---
    if(DOM.btnCreateDailyReport) DOM.btnCreateDailyReport.style.display = 'none';

    function showNotification(message, type='info', duration=3000) { const n=document.createElement('div'); n.className=`notification ${type}`; n.textContent=message; DOM.notificationArea.appendChild(n); n.offsetHeight; n.classList.add('show'); setTimeout(()=>{ n.classList.remove('show'); n.addEventListener('transitionend', ()=>n.remove()); }, duration); }
    function showLoading() { DOM.loadingIndicator.style.display = 'block'; }
    function hideLoading() { DOM.loadingIndicator.style.display = 'none'; }
    function formatNumberFull(num) { num=parseFloat(num); if(isNaN(num)||num===0) return "0"; return parseInt(num).toLocaleString("vi-VN"); }
    function formatNumberHiddenK(num) { num=parseFloat(num); if(isNaN(num)||num===0||Math.abs(num)<1000) return "0"; return parseInt(num/1000).toLocaleString("vi-VN"); }
    function formatCommissionCellHiddenK(rev, com) { return formatNumberHiddenK(com); }
    
    function removeVietnameseTones(str) {
        str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g,"a"); 
        str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g,"e"); 
        str = str.replace(/ì|í|ị|ỉ|ĩ/g,"i"); 
        str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g,"o"); 
        str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g,"u"); 
        str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ/g,"y"); 
        str = str.replace(/đ/g,"d");
        str = str.replace(/À|Á|Ạ|Ả|Ã|Â|Ầ|Ấ|Ậ|Ẩ|Ẫ|Ă|Ằ|Ắ|Ặ|Ẳ|Ẵ/g, "A");
        str = str.replace(/È|É|Ẹ|Ẻ|Ẽ|Ê|Ề|Ế|Ệ|Ể|Ễ/g, "E");
        str = str.replace(/Ì|Í|Ị|Ỉ|Ĩ/g, "I");
        str = str.replace(/Ò|Ó|Ọ|Ỏ|Õ|Ô|Ồ|Ố|Ộ|Ổ|Ỗ|Ơ|Ờ|Ớ|Ợ|Ở|Ỡ/g, "O");
        str = str.replace(/Ù|Ú|Ụ|Ủ|Ũ|Ư|Ừ|Ứ|Ự|Ử|Ữ/g, "U");
        str = str.replace(/Ỳ|Ý|Ỵ|Ỷ|Ỹ/g, "Y");
        str = str.replace(/Đ/g, "D");
        return str;
    }

    function docso(numberStr) {
        const num = String(numberStr).replace(/[.,]/g, ''); if(num==='0') return 'Không'; if(num===''||isNaN(parseInt(num))) return '';
        const words=["không","một","hai","ba","bốn","năm","sáu","bảy","tám","chín"], units=["","nghìn","triệu","tỷ"];
        const readChunk=(s, first)=>{
            if(!s||parseInt(s,10)===0) return ""; let n=s.padStart(3,'0').split('').map(Number);
            if(s.length<3 && first) { if(s.length===1) n=[0,0,Number(s)]; else n=[0,Number(s[0]),Number(s[1])]; }
            const [tr,ch,dv]=n; let str="";
            if(tr>0) str+=words[tr]+" trăm "; else if(!first&&(ch>0||dv>0)) str+="không trăm ";
            if(ch===0&&dv!==0) { if(tr>0||!first) str+="linh "; } else if(ch===1) str+="mười "; else if(ch>1) str+=words[ch]+" mươi ";
            if(dv>0) { if(ch===0) { if((tr>0||!first)&&parseInt(s,10)>0) str+=words[dv]; else if(first) str+=words[dv]; } else if(ch===1) { if(dv===1) str+="một"; else if(dv===5) str+="lăm"; else str+=words[dv]; } else { if(dv===1) str+="mốt"; else if(dv===5) str+="lăm"; else str+=words[dv]; } }
            return str.trim();
        };
        const chunks=[]; let tmp=num; while(tmp.length>3){ chunks.unshift(tmp.substring(Math.max(0,tmp.length-3))); tmp=tmp.substring(0,tmp.length-3); } chunks.unshift(tmp);
        if(chunks.length>units.length) return "Số quá lớn";
        let res=chunks.map((c,i)=>{ const t=readChunk(c,i===0); return t?t+" "+units[chunks.length-1-i]:""; }).filter(Boolean).join(" ").replace(/\s+/g,' ').trim();
        return res.charAt(0).toUpperCase()+res.slice(1);
    }

    async function copyElementAsImage(sel, msg, btn) { 
        const el = document.querySelector(sel); 
        if(!el || !el.innerHTML.trim()) return showNotification("Chưa có nội dung!", 'info');
        
        const oldTxt = btn ? btn.textContent : ''; 
        if(btn) btn.textContent = 'Đang xử lý...';
        
        const hiddenEls = [
            el.querySelector('.summary-actions'), 
            el.querySelector('.controls-panel'), 
            document.getElementById('actionButtonContainer'),
            document.getElementById('sellTicketControls'),
            el.querySelector('.nav-header'),
            document.querySelector('.home-menu-grid')
        ].filter(Boolean);
        hiddenEls.forEach(e => e.style.display = 'none'); 

        const footer = document.createElement('div');
        footer.innerText = "© 2025 ZENG FINANCE";
        footer.style.width = "100%"; footer.style.textAlign = "center"; footer.style.marginTop = "20px"; footer.style.padding = "10px 0"; footer.style.fontSize = "14px"; footer.style.fontWeight = "bold"; footer.style.color = "#7f8c8d"; footer.style.backgroundColor = "#ffffff";
        el.appendChild(footer);

        let originalOverflow = '';
        if(sel === '#summarySection'){ 
            const table = DOM.summaryTableContainer.querySelector('table');
            if(table){
                    originalOverflow = DOM.summaryTableContainer.style.overflowX;
                    DOM.summaryTableContainer.style.width = `${table.scrollWidth + 20}px`;
                    DOM.summaryTableContainer.style.overflowX = 'visible';
                    DOM.summarySection.style.width = `${table.scrollWidth + 40}px`;
                    DOM.mainPage.style.maxWidth = 'none';
            }
        }
        
        let originalReportWidth = '';
        if(sel === '#reportContainer') {
             const table = el.querySelector('table');
             if(table && table.scrollWidth > el.clientWidth) {
                 originalReportWidth = el.style.width;
                 el.style.width = `${table.scrollWidth + 20}px`;
             }
        }

        setTimeout(async()=>{
            try {
                const canvas = await html2canvas(el, { backgroundColor: "#ffffff", scale: 2, useCORS: true, scrollY: -window.scrollY });
                canvas.toBlob(async(blob)=>{
                    try { await navigator.clipboard.write([new ClipboardItem({'image/png':blob})]); showNotification(msg, 'success'); }
                    catch(err) { console.error(err); DOM.imagePreviewHolder.src = canvas.toDataURL(); DOM.imagePreviewModal.style.display = 'block'; showNotification('Copy thủ công từ ảnh dưới.', 'info'); }
                });
            } catch(e){ console.error(e); showNotification('Lỗi tạo ảnh', 'error'); }
            finally {
                if(footer && footer.parentNode) footer.parentNode.removeChild(footer);
                hiddenEls.forEach(e => e.style.display = ''); 
                if(btn) btn.textContent = oldTxt;
                if(sel === '#summarySection'){ DOM.summaryTableContainer.style.width = ''; DOM.summaryTableContainer.style.overflowX = originalOverflow; DOM.summarySection.style.width = ''; DOM.mainPage.style.maxWidth = ''; }
                if(sel === '#reportContainer' && originalReportWidth !== '') { el.style.width = originalReportWidth; }
            }
        }, 100);
    }

    function handleFile(file, num, inputId) {
        if(!file) return;
        showLoading(); const reader=new FileReader();
        reader.onload=(e)=>{
            try {
                const data=new Uint8Array(e.target.result); const wb=XLSX.read(data,{type:"array"}); const ws=wb.Sheets[wb.SheetNames[0]]; const json=XLSX.utils.sheet_to_json(ws,{defval:""});
                if(!json.length) throw new Error("File rỗng");
                
                if(num===1) { 
                    excelData1=json; 
                    let dr=json.find(r=>Object.values(r).some(v=>typeof v==="string" && (v.includes("Tõ ngµy") || v.includes("Ngµy"))));
                    if(dr) { 
                        let txt=Object.values(dr).find(v=>typeof v==="string" && (v.includes("Tõ ngµy") || v.includes("Ngµy")));
                        let mRange=txt.match(/Tõ ngµy\s*(\d{1,2}\/\d{1,2})\/\d{4}.*®Õn ngµy\s*(\d{1,2}\/\d{1,2})\/\d{4}/);
                        let mSingle=txt.match(/Ngµy\s*(\d{1,2}\/\d{1,2}\/\d{4})/);
                        if(mRange) { dateRangeText=(mRange[1]===mRange[2]?`Ngày ${mRange[1]}`:`Từ ${mRange[1]} đến ${mRange[2]}`); } else if (mSingle) { dateRangeText = `Ngày ${mSingle[1]}`; }
                    }
                } else excelData2=json;
                document.getElementById(inputId).classList.add('loaded'); showNotification("Đã nạp file thành công!",'success');
            } catch(err) { console.error(err); showNotification("Lỗi đọc file: "+err.message,'error'); } finally { hideLoading(); }
        }; reader.readAsArrayBuffer(file);
    }

    function parseRawDataInput() {
        const text = DOM.rawInput.value.trim();
        if(!text) return false;
        
        rawDataMap = new Map();
        rawDateRange = "Chưa có ngày";

        const lines = text.split('\n');
        const dateRegex = /Từ ngày\s+(\d{2}\/\d{2}\/\d{4})\s+đến ngày\s+(\d{2}\/\d{2}\/\d{4})/i;
        const csvRegex = /^"([^"]+)","([^"]+)","([^"]+)","([^"]+)","([^"]+)","([^"]+)","([^"]+)"/;

        let count = 0;
        lines.forEach(line => {
            line = line.trim();
            if(!line) return;
            const dateMatch = line.match(dateRegex);
            if(dateMatch) { rawDateRange = `Từ ${dateMatch[1].slice(0,5)} đến ${dateMatch[2].slice(0,5)}`; return; }
            const match = line.match(csvRegex);
            if(match) {
                const code = match[1];
                if(code === "Mã ĐL") return;
                const parseVNNum = (str) => { let s = str.replace(/\./g, ''); s = s.replace(/,/g, '.'); return parseFloat(s) * 1000; };
                const item = { date: match[2], type: match[3], rev: parseVNNum(match[4]), com: parseVNNum(match[5]), bonus: parseVNNum(match[6]), net: parseVNNum(match[7]) };
                if(!rawDataMap.has(code)) rawDataMap.set(code, []);
                rawDataMap.get(code).push(item);
                count++;
            }
        });
        if(count > 0) { dateRangeText = rawDateRange; showNotification(`Đã nạp ${count} dòng dữ liệu thô!`, 'success'); return true; } 
        else { showNotification("Không tìm thấy dữ liệu hợp lệ!", 'error'); return false; }
    }

    function parseDebtData() {
        const input=DOM.debtInput.value.trim(); debtData=new Map(); isDebtFromExcel=false;
        if(!input) { DOM.debtStatusMsg.style.display='none'; generateSummaryView(); return; }
        const lines=input.split('\n'); let count=0;
        lines.forEach(l=>{
            l=l.trim(); if(!l) return; const parts=l.split(/[+\/,:;=.\s]+/).filter(p=>p.trim()!=='');
            if(parts.length>=2) {
                const code=parts[0].trim(), amtRaw=parseFloat(parts[1].trim().replace(/[.,]/g,''));
                if(code && !isNaN(amtRaw)) { debtData.set(code, amtRaw*1000); count++; }
            }
        });
        DOM.debtStatusMsg.textContent=count>0?`Đã nhận: ${count} đại lý nợ.`:''; DOM.debtStatusMsg.style.display=count>0?'block':'none';
        generateSummaryView();
    }
    
    function handleDebtFile(file) {
        if(!file) return; showLoading(); const r=new FileReader();
        r.onload=(e)=>{
            try {
                const d=new Uint8Array(e.target.result), wb=XLSX.read(d,{type:"array"}), ws=wb.Sheets[wb.SheetNames[0]], json=XLSX.utils.sheet_to_json(ws,{header:1});
                debtData=new Map(); isDebtFromExcel=true; let count=0;
                json.forEach(row=>{
                    if(row.length>=2) {
                        const c=String(row[0]).trim(); if(c.toLowerCase().includes("mã")||c==="") return;
                        const amt=parseFloat(String(row[1]).replace(/[.,]/g,''));
                        if(!isNaN(amt)) { debtData.set(c, amt*1000); count++; }
                    }
                });
                if(count>0) { DOM.debtStatusMsg.textContent=`Excel: ${count} đại lý nợ.`; DOM.debtStatusMsg.style.display='block'; DOM.debtFile.classList.add('loaded'); generateSummaryView(); }
                else { showNotification("File nợ không hợp lệ",'warning'); isDebtFromExcel=false; }
            } catch(err) { console.error(err); showNotification("Lỗi file nợ",'error'); } finally { hideLoading(); }
        }; r.readAsArrayBuffer(file);
    }

    function processSummaryData(regionRules, revFilter) {
        let groups = new Map();
        let totals = {rev:0, com:0, th:0, debt:0, final:0, lotoDB:0, lotoCap:0, c227:0, c323:0, xskt:0};

        if(appMode === 'excel') {
            const map=new Map();
            const init=(c)=>{ if(!map.has(c)) map.set(c,{totRev:0,totCom:0,lotoDB:0,lotoCap:0,c227:0,c323:0,xskt:0,thuong:0,debt:0}); };
            excelData1.forEach(r=>{
                const c=String(r.__EMPTY||'').trim(); if(!c||!/\d/.test(c)) return; init(c); const d=map.get(c);
                const ldb=parseFloat(r.__EMPTY_2)||0, lc=parseFloat(r.__EMPTY_6)||0, c2=parseFloat(r.__EMPTY_8)||0, c3=parseFloat(r.__EMPTY_10)||0, th=parseFloat(r.__EMPTY_14)||0;
                d.lotoDB+=ldb; d.lotoCap+=lc; d.c227+=c2; d.c323+=c3; d.thuong+=th; d.totRev+=ldb+lc+c2+c3; d.totCom+=(ldb*0.08)+(lc*0.1)+(c2*0.1)+(c3*0.1);
            });
            excelData2.forEach(r=>{
                const c=String(r.__EMPTY||r.__EMPTY_1||'').trim(); if(!c||!/\d/.test(c)) return; init(c); const d=map.get(c);
                const x=(parseFloat(r.__EMPTY_7)||0)*1000; d.xskt+=x; d.totRev+=x; d.totCom+=x*0.1;
            });
            debtData.forEach((v,k)=>{ init(k); map.get(k).debt=v; });
            const keys=[...map.keys()]; keys.forEach(k=>{ const d=map.get(k); if(d.totRev===0 && d.debt===0 && d.thuong===0) map.delete(k); });
            
            let filteredMap=new Map();
            if(regionRules) map.forEach((v,k)=>{ const numMatch=String(k).match(/\d+$/); if(numMatch){ const n=parseInt(numMatch[0],10); if(regionRules.some(r=>n>=r.min && n<=r.max)) filteredMap.set(k,v); } }); else filteredMap=map;
            
            const sortedKeys=[...filteredMap.keys()].sort((a,b)=>{ const na=parseInt(String(a).match(/\d+$/)?.[0]||0), nb=parseInt(String(b).match(/\d+$/)?.[0]||0); return na-nb; });
            sortedKeys.forEach(k=>{
                const gk=parseInt(String(k).match(/\d+$/)?.[0]||k); if(!groups.has(gk)) groups.set(gk,{agents:[],gt:0,gr:0,gth:0,gdb:0,gx:0});
                const g=groups.get(gk), d=filteredMap.get(k), final=(d.totRev-d.totCom-d.thuong)+d.debt;
                g.agents.push({code:k,data:d});
                g.gt+=final; g.gr+=d.totRev; g.gth+=d.thuong; g.gdb+=d.lotoDB; g.gx+=d.xskt;
            });
            
            const finalGroups=new Map();
            if(revFilter) groups.forEach((v,k)=>{ if(v.gr>=revFilter.min && v.gr<=revFilter.max) finalGroups.set(k,v); }); else return {groups, totals:{}};
            finalGroups.forEach(g=>{ g.agents.forEach(a=>{ 
                const d=a.data; totals.rev+=d.totRev; totals.th+=d.thuong; totals.debt+=d.debt; 
                totals.lotoDB+=d.lotoDB; totals.lotoCap+=d.lotoCap; totals.c227+=d.c227; totals.c323+=d.c323; totals.xskt+=d.xskt;
                totals.final+=(d.totRev-d.totCom-d.thuong)+d.debt;
            });});
            return {groups:finalGroups, totals};

        } else {
            const map = new Map();
            const init=(c)=>{ if(!map.has(c)) map.set(c,{totRev:0,totCom:0,lotoDB:0,lotoCap:0,c227:0,c323:0,xskt:0,thuong:0,debt:0}); };
            
            rawDataMap.forEach((items, code) => {
                init(code); const d = map.get(code);
                items.forEach(item => { d.totRev += item.rev; d.totCom += item.com; d.thuong += item.bonus; d.lotoDB += item.rev; });
            });
            debtData.forEach((v,k)=>{ init(k); map.get(k).debt=v; });

            let filteredMap=new Map();
            if(regionRules) map.forEach((v,k)=>{ const numMatch=String(k).match(/\d+$/); if(numMatch){ const n=parseInt(numMatch[0],10); if(regionRules.some(r=>n>=r.min && n<=r.max)) filteredMap.set(k,v); } }); else filteredMap=map;
            const sortedKeys=[...filteredMap.keys()].sort((a,b)=>{ const na=parseInt(String(a).match(/\d+$/)?.[0]||0), nb=parseInt(String(b).match(/\d+$/)?.[0]||0); return na-nb; });

            sortedKeys.forEach(k=>{
                const gk=parseInt(String(k).match(/\d+$/)?.[0]||k); if(!groups.has(gk)) groups.set(gk,{agents:[],gt:0,gr:0,gth:0,gdb:0,gx:0});
                const g=groups.get(gk), d=filteredMap.get(k), final=(d.totRev-d.totCom-d.thuong)+d.debt;
                g.agents.push({code:k,data:d});
                g.gt+=final; g.gr+=d.totRev; g.gth+=d.thuong; g.gdb+=d.lotoDB; g.gx+=d.xskt;
            });

             const finalGroups=new Map();
            if(revFilter) groups.forEach((v,k)=>{ if(v.gr>=revFilter.min && v.gr<=revFilter.max) finalGroups.set(k,v); }); else return {groups, totals:{}};
            finalGroups.forEach(g=>{ g.agents.forEach(a=>{ 
                const d=a.data; totals.rev+=d.totRev; totals.th+=d.thuong; totals.debt+=d.debt; totals.com+=d.totCom; 
                totals.lotoDB+=d.lotoDB; totals.final+=(d.totRev-d.totCom-d.thuong)+d.debt;
            });});
            return {groups:finalGroups, totals};
        }
    }

    function generateSummaryView() {
        if(appMode === 'excel' && !excelData1.length && !excelData2.length && debtData.size===0) return;
        if(appMode === 'raw' && rawDataMap.size===0 && debtData.size===0) return;

        const regVal=DOM.regionSelect.value, rules=regVal!=='all'?regionDefinitions[regVal]:null;
        const minV=parseFloat(DOM.minRevenueInput.dataset.value)||0, maxV=parseFloat(DOM.maxRevenueInput.dataset.value)||Infinity;
        
        const {groups, totals} = processSummaryData(rules, {min:minV*1000, max:maxV*1000});
        
        const hasXSKT = (appMode === 'excel' && excelData2.length > 0);
        const hasDebt = debtData.size > 0;

        let html;
        if (appMode === 'excel') {
             html=`<thead><tr><th width="8%">STT</th><th width="12%">ĐL</th><th>LTô</th><th>LCặp</th><th>Lô 3/23</th><th>Lô 2/27</th>${hasXSKT?'<th>XSKT':''}<th>Dsố</th><th>Thưởng</th>${hasDebt?'<th>Nợ cũ':''}<th>Tiền nộp</th></tr></thead><tbody>`;
        } else {
             html=`<thead><tr><th width="8%">STT</th><th width="12%">ĐL</th><th>Doanh Số</th><th>Hoa hồng</th><th>Thưởng</th>${hasDebt?'<th>Nợ cũ':''}<th>Tiền nộp</th></tr></thead><tbody>`;
        }

        let stt=1, sortedGroups=[...groups.entries()];
        const sortMode=DOM.sortSelect.value;
        if(sortMode==='revenue_desc') sortedGroups.sort((a,b)=>b[1].gr - a[1].gr);
        else if(sortMode==='revenue_asc') sortedGroups.sort((a,b)=>a[1].gr - b[1].gr);

        sortedGroups.forEach(([gk, g])=>{
            g.agents.forEach((ag, idx)=>{
                const d=ag.data, final=(d.totRev-d.totCom-d.thuong)+d.debt;
                if (appMode === 'excel') {
                    html+=`<tr${idx===0?' class="group-start-row"':''}><td>${idx===0?stt:''}</td><td>${ag.code}</td><td>${formatNumberHiddenK(d.lotoDB)}</td><td>${formatNumberHiddenK(d.lotoCap)}</td><td>${formatNumberHiddenK(d.c323)}</td><td>${formatNumberHiddenK(d.c227)}</td>${hasXSKT?`<td>${formatNumberHiddenK(d.xskt)}</td>`:''}<td>${formatNumberHiddenK(d.totRev)}</td><td>${formatNumberHiddenK(d.thuong)}</td>${hasDebt?`<td>${formatNumberHiddenK(d.debt)}</td>`:''}<td>${final<0?`<span class="negative-value">`:''}${formatNumberHiddenK(final)}${final<0?'</span>':''}</td></tr>`;
                } else {
                    html+=`<tr${idx===0?' class="group-start-row"':''}><td>${idx===0?stt:''}</td><td>${ag.code}</td><td>${formatNumberHiddenK(d.totRev)}</td><td>${formatCommissionCellHiddenK(d.totRev, d.totCom)}</td><td>${formatNumberHiddenK(d.thuong)}</td>${hasDebt?`<td>${formatNumberHiddenK(d.debt)}</td>`:''}<td>${final<0?`<span class="negative-value">`:''}${formatNumberHiddenK(final)}${final<0?'</span>':''}</td></tr>`;
                }
            }); 
            if (g.agents.length > 1) {
                let sumDebt = 0; g.agents.forEach(ag => sumDebt+=ag.data.debt);
                if (appMode === 'excel') {
                    let sumLoto = 0, sumCap = 0, sumC3 = 0, sumC2 = 0, sumXskt = 0;
                    g.agents.forEach(ag => { sumLoto+=ag.data.lotoDB; sumCap+=ag.data.lotoCap; sumC3+=ag.data.c323; sumC2+=ag.data.c227; sumXskt+=ag.data.xskt; });
                    html += `<tr class="group-total-row"><td></td><td></td><td>${formatNumberHiddenK(sumLoto)}</td><td>${formatNumberHiddenK(sumCap)}</td><td>${formatNumberHiddenK(sumC3)}</td><td>${formatNumberHiddenK(sumC2)}</td>${hasXSKT?`<td>${formatNumberHiddenK(sumXskt)}</td>`:''}<td>${formatNumberHiddenK(g.gr)}</td><td>${formatNumberHiddenK(g.gth)}</td>${hasDebt?`<td>${formatNumberHiddenK(sumDebt)}</td>`:''}<td style="color:#c0392b;">${formatNumberHiddenK(g.gt)}</td></tr>`;
                } else {
                    html += `<tr class="group-total-row"><td></td><td></td><td>${formatNumberHiddenK(g.gr)}</td><td></td><td>${formatNumberHiddenK(g.gth)}</td>${hasDebt?`<td>${formatNumberHiddenK(sumDebt)}</td>`:''}<td style="color:#c0392b;">${formatNumberHiddenK(g.gt)}</td></tr>`;
                }
            }
            stt++;
        });

        if(appMode === 'excel') {
            html+=`</tbody><tfoot><tr><td colspan="2">TỔNG</td><td>${formatNumberHiddenK(totals.lotoDB)}</td><td>${formatNumberHiddenK(totals.lotoCap)}</td><td>${formatNumberHiddenK(totals.c323)}</td><td>${formatNumberHiddenK(totals.c227)}</td>${hasXSKT?`<td>${formatNumberHiddenK(totals.xskt)}</td>`:''}<td>${formatNumberHiddenK(totals.rev)}</td><td>${formatNumberHiddenK(totals.th)}</td>${hasDebt?`<td>${formatNumberHiddenK(totals.debt)}</td>`:''}<td class="${totals.final>=0?'final-amount-positive':'final-amount-negative'}">${formatNumberHiddenK(totals.final)}</td></tr></tfoot>`;
        } else {
             html+=`</tbody><tfoot><tr><td colspan="2">TỔNG</td><td>${formatNumberHiddenK(totals.rev)}</td><td>${formatNumberHiddenK(totals.com)}</td><td>${formatNumberHiddenK(totals.th)}</td>${hasDebt?`<td>${formatNumberHiddenK(totals.debt)}</td>`:''}<td class="${totals.final>=0?'final-amount-positive':'final-amount-negative'}">${formatNumberHiddenK(totals.final)}</td></tr></tfoot>`;
        }

        DOM.summaryTableContainer.innerHTML=`<table border="1">${html}</table>`;
        const displayRegion = regVal === 'all' ? 'Tất Cả' : regVal;
        DOM.summaryTitle.innerHTML=`Bảng Tổng Hợp: <span class="region-name">${displayRegion}</span>`;
        let percent = totals.rev > 0 ? (totals.th / totals.rev * 100) : 0;
        let pString = `<span style="color:${percent<=50?'var(--success-color)':'var(--danger-color)'}">(${percent.toFixed(1)}%)</span>`;
        DOM.summaryDate.innerHTML=`${dateRangeText}<br>Doanh số: <b style="color:var(--danger-color)">${formatNumberHiddenK(totals.rev)}</b> - Thưởng: <b style="color:var(--primary-dark)">${formatNumberHiddenK(totals.th)}</b> ${pString}`;
    }

    function resetState(soSo) { state.currentSoSo=soSo; state.baseTongDT=0; state.baseTongHH=0; state.baseThuong=0; state.extraBonus=0; state.extraDebt=0; state.addedTickets=[]; state.accountCodes=new Set(); state.debtFromData=0; }
    
    function processDetailedReport(soSo) {
        if(appMode === 'excel') {
            const targetNum = parseInt(soSo, 10);
            if (isNaN(targetNum)) return null; 
            const filter = (r) => { const c = String(r.__EMPTY || r.__EMPTY_1 || '').trim(); const match = c.match(/(\d+)$/); return match && parseInt(match[1], 10) === targetNum; };
            const f1 = excelData1.filter(filter); const f2 = excelData2.filter(filter);
            
            let debtAmt = 0; debtData.forEach((val, key) => { const match = String(key).match(/(\d+)$/); if(match && parseInt(match[1], 10) === targetNum) { debtAmt += val; state.accountCodes.add(key); } });
            if (!f1.length && !f2.length && debtAmt === 0) return null;

            resetState(soSo);
            const t = {lotoDB: 0, loCap: 0, c227: 0, c323: 0, thuong: 0, xskt: 0};
            f1.forEach(r => { state.accountCodes.add(String(r.__EMPTY).trim()); t.lotoDB += (parseFloat(r.__EMPTY_2) || 0); t.loCap += (parseFloat(r.__EMPTY_6) || 0); t.c227 += (parseFloat(r.__EMPTY_8) || 0); t.c323 += (parseFloat(r.__EMPTY_10) || 0); t.thuong += (parseFloat(r.__EMPTY_14) || 0); });
            f2.forEach(r => { state.accountCodes.add(String(r.__EMPTY || r.__EMPTY_1).trim()); t.xskt += (parseFloat(r.__EMPTY_7) || 0) * 1000; });
            const c = {ldb: t.lotoDB * 0.08, lc: t.loCap * 0.1, c2: t.c227 * 0.1, c3: t.c323 * 0.1, x: t.xskt * 0.1};
            state.baseTongDT = t.lotoDB + t.loCap + t.c227 + t.c323 + t.xskt; state.baseTongHH = c.ldb + c.lc + c.c2 + c.c3 + c.x; state.baseTongHH = c.ldb + c.lc + c.c2 + c.c3 + c.x; state.baseThuong = t.thuong; state.debtFromData = debtAmt;
            return {t, c};
        } else {
            // === LOGIC MỚI: TÌM KIẾM TƯƠNG ĐỐI ===
            let targetKey = null;
            const inputNum = soSo.replace(/\D/g, ''); 
            for (let key of rawDataMap.keys()) {
                const keyNum = key.replace(/\D/g, '');
                if (key === soSo || (inputNum.length > 0 && keyNum === inputNum)) {
                    targetKey = key;
                    break;
                }
            }
            if(!targetKey) return null;
            
            const items = rawDataMap.get(targetKey);
            let debtAmt = 0; if(debtData.has(targetKey)) debtAmt = debtData.get(targetKey);
            
            resetState(targetKey);
            state.accountCodes.add(targetKey);
            
            let totalRev = 0, totalCom = 0, totalBonus = 0;
            items.forEach(i => { totalRev += i.rev; totalCom += i.com; totalBonus += i.bonus; });
            
            state.baseTongDT = totalRev; state.baseTongHH = totalCom; state.baseThuong = totalBonus; state.debtFromData = debtAmt;
            return items; 
        }
    }

    // --- CẬP NHẬT RENDER DETAILED REPORT ---
    function renderDetailedReport(d, customDate, reportType = 'detailed') {
        const debtHtml = state.debtFromData > 0 ? `<tr style="color:var(--danger-color); font-weight:bold;"><td>Nợ cũ</td><td colspan="${(appMode === 'excel' || (appMode === 'raw' && reportType === 'summary')) ? 2 : 4}">${formatNumberHiddenK(state.debtFromData)}</td></tr>` : '';
        const reportDate = customDate || dateRangeText;

        // LOGIC MỚI: Sử dụng giao diện Excel (Tổng hợp) nếu:
        // 1. Đang ở chế độ Excel
        // 2. HOẶC Đang ở chế độ Raw nhưng người dùng chọn "Báo Cáo Tổng"
        if (appMode === 'excel' || (appMode === 'raw' && reportType === 'summary')) {
            let t, c;

            if (appMode === 'excel') {
                // Dữ liệu từ Excel đã có sẵn cấu trúc t, c
                t = d.t; 
                c = d.c;
            } else {
                // CHUYỂN ĐỔI DỮ LIỆU THÔ (ARRAY) -> DẠNG TỔNG HỢP (OBJECT t, c)
                t = {lotoDB: 0, loCap: 0, c227: 0, c323: 0, thuong: 0, xskt: 0};
                c = {ldb: 0, lc: 0, c2: 0, c3: 0, x: 0};

                // Duyệt qua mảng dữ liệu thô để cộng dồn
                d.forEach(item => {
                    t.thuong += item.bonus;
                    
                    // Map loại hình thô sang biến tổng
                    const type = item.type.toUpperCase();
                    if (type.includes('LT') || type.includes('DB')) {
                        t.lotoDB += item.rev; c.ldb += item.com;
                    } else if (type.includes('LC') || type.includes('CAP')) {
                        t.loCap += item.rev; c.lc += item.com;
                    } else if (type.includes('L2') || type.includes('2/27')) {
                        t.c227 += item.rev; c.c2 += item.com;
                    } else if (type.includes('L3') || type.includes('3/23')) {
                        t.c323 += item.rev; c.c3 += item.com;
                    }
                    // Nếu có loại khác, nó sẽ được tính vào tổng doanh thu chung ở hàm updateDetailedView
                });
            }

            const xsktHtml = t.xskt > 0 ? `<tr><td>XSKT</td><td>${formatNumberHiddenK(t.xskt)}</td><td>${formatCommissionCellHiddenK(t.xskt, c.x)}</td></tr>` : '';
            DOM.reportContainer.innerHTML = `
            <table class="report-table" style="border: 1px solid var(--border-color);">
                <colgroup><col style="width: 30%;"><col style="width: 35%;"><col style="width: 35%;"></colgroup>
                <tbody>
                    <tr><td class="header" style="font-weight:bold;">Đại Lý:</td><td colspan="2" class="so-so" style="text-align:center;">${[...state.accountCodes].join(', ')}</td></tr>
                    <tr><td class="header" style="font-weight:bold;">Ngày:</td><td colspan="2" style="font-weight:500;">${reportDate}</td></tr>
                    <tr style="background:#f1f2f6"><th>Loại hình</th><th>Doanh thu</th><th>Hoa hồng</th></tr>
                    <tr><td>Lôtô ĐB</td><td>${formatNumberHiddenK(t.lotoDB)}</td><td>${formatCommissionCellHiddenK(t.lotoDB, c.ldb)}</td></tr>
                    <tr><td>Lôtô cặp</td><td>${formatNumberHiddenK(t.loCap)}</td><td>${formatCommissionCellHiddenK(t.loCap, c.lc)}</td></tr>
                    <tr><td>Lô 2/27</td><td>${formatNumberHiddenK(t.c227)}</td><td>${formatCommissionCellHiddenK(t.c227, c.c2)}</td></tr>
                    <tr><td>Lô 3/23</td><td>${formatNumberHiddenK(t.c323)}</td><td>${formatCommissionCellHiddenK(t.c323, c.c3)}</td></tr>
                    ${xsktHtml}
                    <tr id="tongCongRow" style="border-top: 2px solid #bdc3c7;"><td>Tổng cộng</td><td id="tongDT"></td><td id="tongHH"></td></tr>
                    <tr id="mainBonusRow"><td>Thưởng</td><td colspan="2">${formatNumberHiddenK(state.baseThuong)}</td></tr>
                    ${debtHtml}
                    <tr id="finalResultRow"><td id="ketQuaLabel"></td><td colspan="2" id="ketQuaValue"></td></tr>
                    <tr id="bangChuRow"><td class="header" style="font-size:13px;">Bằng chữ</td><td colspan="2" id="ketQuaBangChu" class="bang-chu-cell"></td></tr>
                </tbody>
            </table>`;
        } else {
            // === LOGIC CŨ: RAW DATA REPORT CHI TIẾT ===
            
            // 1. Nhóm dữ liệu theo ngày
            const groupByDate = new Map();
            d.forEach(item => {
                if(!groupByDate.has(item.date)) groupByDate.set(item.date, []);
                groupByDate.get(item.date).push(item);
            });

            let rowsHtml = '';
            
            // Map đổi tên loại hình
            const typeNameMap = {
                'LT': 'Ltô ĐB',
                'LC': 'LCặp',
                'L2': 'Lô 2/27',
                'L3': 'Lô 3/23'
            };

            // 2. Duyệt từng ngày để hiển thị
            groupByDate.forEach((items, dateStr) => {
                const isSingleItem = items.length === 1; // Kiểm tra nếu chỉ có 1 loại hình
                let dayRev = 0, dayCom = 0, dayBonus = 0;
                
                // Duyệt item trong ngày
                items.forEach(item => {
                    dayRev += item.rev;
                    dayCom += item.com;
                    dayBonus += item.bonus;
                    
                    const displayName = typeNameMap[item.type] || item.type;

                    // Xác định style cho dòng item này
                    let cellStyle = '';
                    let revColor = '';
                    let netColor = 'font-weight:bold; color: #000;'; // Mặc định đen cho item lẻ trong ngày nhiều loại

                    if (isSingleItem) {
                        // NẾU LÀ NGÀY ĐƠN: Áp dụng style của dòng tổng
                        cellStyle = 'background-color: #ecf0f1;'; // Nền xám
                        revColor = 'color: #000;'; // Doanh thu đen
                        // Nộp/Lấy về màu Đỏ/Xanh
                        if (item.net >= 0) netColor = 'font-weight:bold; color: var(--danger-color);'; 
                        else netColor = 'font-weight:bold; color: var(--success-color);';
                    }

                    rowsHtml += `<tr>
                        <td style="text-align:left; font-size:14px;">${item.date} - <b style="color:#2c3e50">${displayName}</b></td>
                        <td style="${revColor} ${cellStyle}">${formatNumberHiddenK(item.rev)}</td>
                        <td style="${cellStyle}">${formatNumberHiddenK(item.com)}</td>
                        <td style="${cellStyle}">${item.bonus > 0 ? formatNumberHiddenK(item.bonus) : '-'}</td>
                        <td style="${netColor} ${cellStyle}">${formatNumberHiddenK(item.net)}</td>
                    </tr>`;
                });

                // 3. Chỉ hiện dòng TỔNG NGÀY nếu có > 1 loại hình
                if (!isSingleItem) {
                    const dayNet = dayRev - dayCom - dayBonus;
                    const cellBgStyle = 'background-color: #ecf0f1;'; 
                    const netColorTotal = dayNet >= 0 ? 'var(--danger-color)' : 'var(--success-color)';

                    rowsHtml += `<tr style="font-weight: bold; border-bottom: 2px solid #bdc3c7;">
                        <td style="text-align:right; font-style: italic;"></td> 
                        <td style="color: #000; ${cellBgStyle}">${formatNumberHiddenK(dayRev)}</td>
                        <td style="color: #000; ${cellBgStyle}">${formatNumberHiddenK(dayCom)}</td>
                        <td style="color: #000; ${cellBgStyle}">${formatNumberHiddenK(dayBonus)}</td>
                        <td style="color: ${netColorTotal}; ${cellBgStyle}">${formatNumberHiddenK(dayNet)}</td>
                    </tr>`;
                }
            });

            // 3. Tổng hợp loại hình
            let typeSummaryHtml = '';
            const sortOrder = ['LT', 'LC', 'L2', 'L3'];

            // Gom nhóm tổng theo loại hình (từ dữ liệu đầu vào d)
            const typeSummary = new Map();
            d.forEach(item => {
                if(!typeSummary.has(item.type)) typeSummary.set(item.type, {rev:0, com:0, bonus:0});
                const t = typeSummary.get(item.type);
                t.rev += item.rev;
                t.com += item.com;
                t.bonus += item.bonus; 
            });

            if (typeSummary.size > 0) {
                 typeSummaryHtml += `<tr style="border-top: 2px solid #bdc3c7;"><td colspan="5" style="background:#ecf0f1; height: 5px; padding:0;"></td></tr>`; 
                 
                 const currentTypes = Array.from(typeSummary.keys());
                 currentTypes.sort((a, b) => {
                     let idxA = sortOrder.indexOf(a); let idxB = sortOrder.indexOf(b);
                     if (idxA === -1) idxA = 999; if (idxB === -1) idxB = 999;
                     return idxA - idxB;
                 });

                 currentTypes.forEach(type => {
                     const val = typeSummary.get(type);
                     const longName = typeNameMap[type] || type; 
                     const displayName = typeNameMap[type] ? `Tổng ${longName}` : `Tổng ${type}`;

                     typeSummaryHtml += `<tr style="background-color: #fff;">
                        <td style="font-weight:600; color:#5D4037; text-align: left; padding-left: 15px;">${displayName}</td>
                        <td style="font-weight:600;">${formatNumberHiddenK(val.rev)}</td>
                        <td>${formatCommissionCellHiddenK(val.rev, val.com)}</td>
                        <td style="color:#2c3e50; font-weight:500;">${val.bonus > 0 ? formatNumberHiddenK(val.bonus) : '-'}</td>
                        <td></td>
                     </tr>`;
                 });
            }

            DOM.reportContainer.innerHTML = `
            <table class="report-table" style="border: 1px solid var(--border-color);">
                <colgroup><col style="width: 25%;"><col style="width: 20%;"><col style="width: 15%;"><col style="width: 15%;"><col style="width: 25%;"></colgroup>
                <tbody>
                    <tr><td class="header" style="font-weight:bold;">Đại Lý:</td><td colspan="4" class="so-so" style="text-align:center;">${state.currentSoSo}</td></tr>
                    <tr><td class="header" style="font-weight:bold;">Đợt:</td><td colspan="4" style="font-weight:500;">${reportDate}</td></tr>
                    <tr><th>Ngày/Loại</th><th>Doanh Thu</th><th>Hoa hồng</th><th>Thưởng</th><th>Nộp/Lấy về</th></tr>
                    ${rowsHtml}
                    
                    ${typeSummaryHtml}

                    <tr id="tongCongRow" style="border-top: 2px solid #7f8c8d; background-color: #FEF9E7;">
                        <td style="font-weight: bold; color: var(--header-bg);">TỔNG CỘNG</td>
                        <td id="tongDT" style="font-weight:bold; color:var(--danger-color); font-size:1.1em;"></td>
                        <td id="tongHH" style="font-weight:bold;"></td>
                        <td style="font-weight:bold; color:var(--primary-dark)">${formatNumberHiddenK(state.baseThuong)}</td>
                        <td></td>
                    </tr>
                    ${debtHtml}
                    <tr id="finalResultRow"><td id="ketQuaLabel"></td><td colspan="4" id="ketQuaValue"></td></tr>
                    <tr id="bangChuRow"><td class="header" style="font-size:13px;">Bằng chữ</td><td colspan="4" id="ketQuaBangChu" class="bang-chu-cell"></td></tr>
                </tbody>
            </table>`;
        }
        const cpy = document.getElementById('copyrightInReport'); if(!cpy) { const dv=document.createElement('div'); dv.id='copyrightInReport'; dv.className='copyright'; dv.style.display='none'; dv.textContent='© 2025 ZENG FINANCE'; DOM.reportContainer.appendChild(dv); }
    }

    function updateDetailedView() {
        if(!state.currentSoSo) return;
        const addRev=state.addedTickets.reduce((s,t)=>s+t.dt,0), addCom=state.addedTickets.reduce((s,t)=>s+t.hh,0);
        const totDT=state.baseTongDT+addRev, totHH=state.baseTongHH+addCom;
        const final = (totDT - totHH - state.baseThuong) + state.debtFromData - state.extraBonus + state.extraDebt;
        
        document.getElementById('tongDT').textContent=formatNumberHiddenK(totDT); 
        document.getElementById('tongHH').textContent=formatNumberHiddenK(totHH);
        
        const rRow=document.getElementById('finalResultRow'), lbl=document.getElementById('ketQuaLabel'), val=document.getElementById('ketQuaValue'), txt=document.getElementById('ketQuaBangChu');
        rRow.className=final>=0?'final-amount-positive':'final-amount-negative';
        
        if(final>=0) { const round=Math.ceil(final/1000)*1000; lbl.textContent='Nộp'; val.textContent=formatNumberHiddenK(round); txt.textContent=docso(round)+" đồng"; } 
        else { const abs=Math.abs(final), round=parseInt(abs/1000)*1000; lbl.textContent='Lấy về'; val.textContent=formatNumberHiddenK(round); txt.textContent=docso(round)+" đồng"; }
    }

    function createReport(type = 'detailed') {
        DOM.manualInputModal.style.display='none'; 
        const s=DOM.soSoInput.value.trim(); 
        if(!s) return showNotification("Nhập số sổ!",'error');
        
        const d=processDetailedReport(s); 
        if(!d) return showNotification("Không có dữ liệu",'error');
        
        // Truyền tham số type vào hàm render
        renderDetailedReport(d, null, type); 
        
        updateDetailedView(); 
        DOM.actionButtonContainer.style.display='grid'; 
        DOM.actionChoices.style.display='none'; 
        DOM.toggleChoicesBtn.textContent='Thêm tiền Thưởng/Nợ cũ'; 
        DOM.toggleChoicesBtn.className='secondary';
        DOM.reportContainer.scrollIntoView({behavior:'smooth'});
    }

    // --- EVENT LISTENERS CẬP NHẬT: ẨN/HIỆN NÚT THỦ CÔNG ---
    DOM.tabExcel.onclick = () => {
        appMode = 'excel';
        DOM.tabExcel.classList.add('active'); DOM.tabExcel.style.background = 'var(--primary-color)'; DOM.tabExcel.style.color = 'white';
        DOM.tabRaw.classList.remove('active'); DOM.tabRaw.style.background = '#ecf0f1'; DOM.tabRaw.style.color = 'var(--text-color)';
        DOM.panelExcel.style.display = 'block'; DOM.panelRaw.style.display = 'none';
        excelData1 = []; excelData2 = []; rawDataMap = new Map();
        document.querySelectorAll('input[type="file"]').forEach(i => i.classList.remove('loaded'));
        
        // Hiện nút thủ công (Cả ở home và inline)
        if(DOM.btnShowManualReport_inline) DOM.btnShowManualReport_inline.style.display = 'block'; 
        if(DOM.btnShowManualReport) DOM.btnShowManualReport.style.display = 'block';

        // Ẩn nút "Báo Cáo Ngày"
        if(DOM.btnCreateDailyReport) DOM.btnCreateDailyReport.style.display = 'none';
    };

    DOM.tabRaw.onclick = () => {
        appMode = 'raw';
        DOM.tabRaw.classList.add('active'); DOM.tabRaw.style.background = 'var(--primary-color)'; DOM.tabRaw.style.color = 'white';
        DOM.tabExcel.classList.remove('active'); DOM.tabExcel.style.background = '#ecf0f1'; DOM.tabExcel.style.color = 'var(--text-color)';
        DOM.panelRaw.style.display = 'block'; DOM.panelExcel.style.display = 'none';
        excelData1 = []; excelData2 = []; rawDataMap = new Map();
        
        // Ẩn nút thủ công (Cả ở home và inline)
        if(DOM.btnShowManualReport_inline) DOM.btnShowManualReport_inline.style.display = 'none';
        if(DOM.btnShowManualReport) DOM.btnShowManualReport.style.display = 'none';

        // Hiện nút "Báo Cáo Ngày"
        if(DOM.btnCreateDailyReport) DOM.btnCreateDailyReport.style.display = 'block';
    };

    DOM.btnLoadData.addEventListener('click', ()=>{ 
        if(appMode === 'excel') {
            if(!excelData1.length && !excelData2.length) return showNotification("Chọn ít nhất 1 file Excel!",'error');
        } else {
            const success = parseRawDataInput();
            if(!success) return; 
        }
        
        generateSummaryView(); 
        DOM.homePage.style.display='none'; 
        DOM.mainPage.style.display='block'; 
        showSection('reportSection'); 
        DOM.soSoInput.focus(); 
        runLogoAnimation(); 
    });

    function addTicket() {
        if(!state.currentSoSo) return showNotification("Chưa có báo cáo để thêm vé!",'error');
        const q = parseInt(DOM.sellQuantityInput.value, 10); if(!q || q <= 0) return showNotification("Số lượng không hợp lệ!",'error');
        let typeRad; try { typeRad = document.querySelector('input[name="sellTicketType"]:checked').value; } catch(e) { typeRad = 'xo-so'; }
        let name = ""; if(typeRad === 'custom-name') { name = DOM.customTicketName.value.trim(); if(!name) return showNotification("Vui lòng nhập tên loại hình!", 'error'); } else { name = typeRad === 've-boc' ? 'Vé bóc' : 'Xổ số'; }
        let priceRad, price = 0; try { priceRad = document.querySelector('input[name="sellPrice"]:checked').value; } catch(e) { priceRad = '10000'; }
        if(priceRad === 'custom-price') { price = parseFloat(DOM.customTicketPrice.value.replace(/\D/g, '')) || 0; if(price <= 0) return showNotification("Vui lòng nhập giá bán hợp lệ!", 'error'); } else { price = parseFloat(priceRad); }
        let rateRad, ratePercent = 0; try { rateRad = document.querySelector('input[name="sellRate"]:checked').value; } catch(e) { rateRad = '10'; }
        if (rateRad === 'custom-rate') { ratePercent = parseFloat(DOM.customTicketRate.value); if (isNaN(ratePercent) || ratePercent < 0) return showNotification("Vui lòng nhập % hoa hồng hợp lệ!", 'error'); } else { ratePercent = parseFloat(rateRad); }
        const dt = q * price; const hh = dt * (ratePercent / 100); 
        state.addedTickets.push({dt, hh});
        const tr = document.createElement('tr'); tr.className = 'added-ticket-row'; 
        tr.innerHTML = `<td>${name} (${formatNumberHiddenK(price)}k) [${q} vé]</td><td>${formatNumberHiddenK(dt)}</td><td>${formatCommissionCellHiddenK(dt, hh)}</td>`;
        const anchor = document.getElementById('tongCongRow');
        if(anchor) { anchor.parentNode.insertBefore(tr, anchor); updateDetailedView(); DOM.sellQuantityInput.value = ''; showNotification("Đã thêm vé thành công!", 'success'); showSection('reportSection'); } else { showNotification("Lỗi: Không tìm thấy bảng báo cáo!", 'error'); }
    }

    DOM.toggleChoicesBtn.onclick = () => { const show = DOM.toggleChoicesBtn.textContent.includes('Thêm'); DOM.actionChoices.style.display = show ? 'grid' : 'none'; DOM.toggleChoicesBtn.textContent = show ? 'Đóng / Xoá hết' : 'Thêm tiền Thưởng/Nợ cũ'; DOM.toggleChoicesBtn.className = show ? 'danger' : 'secondary'; if(!show) { document.getElementById('extraBonusRow')?.remove(); document.getElementById('extraDebtRow')?.remove(); state.extraBonus=0; state.extraDebt=0; updateDetailedView(); } renderActionButtons(); };
    function renderActionButtons() { DOM.actionChoices.innerHTML = `${document.getElementById('extraBonusRow') ? `<button class="danger" onclick="window.removeRow('bonus')">Xoá Thưởng thêm</button>` : `<button onclick="window.addRow('bonus')">Thêm Thưởng</button>`}${document.getElementById('extraDebtRow') ? `<button class="danger" onclick="window.removeRow('debt')">Xoá Nợ cũ</button>` : `<button onclick="window.addRow('debt')">Thêm Nợ cũ</button>`}`; }
    window.addRow = (type) => { const isB = type === 'bonus', rowId = isB ? 'extraBonusRow' : 'extraDebtRow'; const anchor = document.getElementById('extraBonusRow') || document.getElementById('mainBonusRow'); if(document.getElementById(rowId)) return; const tr = document.createElement('tr'); tr.id = rowId; tr.innerHTML = `<td>${isB ? `<textarea class="dynamic-input" style="color:var(--success-color)" placeholder="Tên thưởng..." rows="1"></textarea>` : 'Nợ cũ'}</td><td colspan="${(appMode==='excel' || appMode==='raw')?2:4}"><input class="dynamic-input" style="color:${isB ? 'var(--success-color)' : 'var(--danger-color)'}" oninput="window.handleDynInp(this,'${type}')" onkeydown="if(event.key==='Enter') this.blur()" inputmode="decimal" placeholder="0"></td>`; anchor.after(tr); renderActionButtons(); const input = tr.querySelector('input'); if(input) input.focus(); };
    window.removeRow = (type) => { document.getElementById(type === 'bonus' ? 'extraBonusRow' : 'extraDebtRow')?.remove(); if(type === 'bonus') state.extraBonus = 0; else state.extraDebt = 0; updateDetailedView(); renderActionButtons(); };
    window.handleDynInp = (el, type) => { const v = parseInt(el.value.replace(/\D/g,'')) || 0; el.value = v.toLocaleString('vi-VN'); if(type === 'bonus') state.extraBonus = v * 1000; else state.extraDebt = v * 1000; updateDetailedView(); };
    function showSection(id) { ['summarySection','reportSection','sellSection'].forEach(s=>DOM[s].style.display='none'); DOM[id].style.display='block'; }
    async function runLogoAnimation() { const containers = document.querySelectorAll('.simple-animated-logo'); if (containers.length === 0) return; const getAllSpans = () => { const all = []; containers.forEach(c => { all.push(Array.from(c.querySelectorAll('span.letter'))); }); return all; }; while (true) { const spanGroups = getAllSpans(); const hue = Math.floor(Math.random() * 360); const color = `hsl(${hue}, 70%, 50%)`; spanGroups.forEach(group => { group.forEach(s => { s.style.transition = 'none'; s.style.opacity = '0'; s.style.transform = 'translateY(20px)'; s.style.color = color; }); }); void containers[0].offsetWidth; const delayStep = 100; const maxIndex = spanGroups[0].length; for (let i = 0; i < maxIndex; i++) { spanGroups.forEach(group => { if (group[i]) { group[i].style.transition = 'all 0.5s cubic-bezier(0.175, 0.885, 0.32, 1.275)'; setTimeout(() => { group[i].style.opacity = '1'; group[i].style.transform = 'translateY(0)'; }, i * delayStep); } }); } const animationTime = (maxIndex * delayStep) + 500; await new Promise(r => setTimeout(r, animationTime + 4000)); for (let i = 0; i < maxIndex; i++) { spanGroups.forEach(group => { if (group[i]) { group[i].style.transition = 'all 0.4s ease-in'; setTimeout(() => { group[i].style.opacity = '0'; group[i].style.transform = 'translateY(-20px)'; }, i * 50); } }); } await new Promise(r => setTimeout(r, 600)); } }
    function updateOnlineCounter() { const num = Math.floor(Math.random() * 98) + 1; document.querySelectorAll('.online-count').forEach(el => el.textContent = num); setTimeout(updateOnlineCounter, Math.random() * 5000 + 3000); }
    function printReport() { if(DOM.summarySection.style.display !== 'none') { window.print(); } else { showNotification("Vui lòng chuyển sang tab Báo cáo Tổng hợp để in!", "warning"); } }
    
    // --- HÀM EXPORT EXCEL ---
    function exportSummaryToExcel() {
        const regVal = DOM.regionSelect.value; const rules = regVal !== 'all' ? regionDefinitions[regVal] : null;
        const minV = parseFloat(DOM.minRevenueInput.dataset.value) || 0; const maxV = parseFloat(DOM.maxRevenueInput.dataset.value) || Infinity;
        const data = processSummaryData(rules, {min: minV * 1000, max: maxV * 1000});
        if (!data) return showNotification("Không có dữ liệu!", 'error');
        
        const {groups, totals} = data; const hasXSKT = excelData2.length > 0 && appMode === 'excel'; const hasDebt = debtData.size > 0;
        
        // --- ĐỊNH NGHĨA STYLES ---
        const sHeader = { font: { bold: true, color: { rgb: "FFFFFF" } }, fill: { fgColor: { rgb: "34495E" } }, alignment: { horizontal: "center", vertical: "center" }, border: { top: {style:"thin"}, bottom: {style:"thin"}, left: {style:"thin"}, right: {style:"thin"} } };
        const sGroup = { font: { bold: true }, fill: { fgColor: { rgb: "FEF9E7" } }, border: { top: {style:"thin"}, bottom: {style:"thin"}, left: {style:"thin"}, right: {style:"thin"} } };
        // Style cho dòng Tổng cuối cùng (Màu vàng đậm hơn chút)
        const sTotal = { font: { bold: true }, fill: { fgColor: { rgb: "FFC107" } }, border: { top: {style:"thin"}, bottom: {style:"thin"}, left: {style:"thin"}, right: {style:"thin"} } };
        
        const sNormal = { border: { top: {style:"thin"}, bottom: {style:"thin"}, left: {style:"thin"}, right: {style:"thin"} } };
        // Style cho STT và ĐL (In đậm + Căn giữa)
        const sCenterBold = { font: { bold: true }, alignment: { horizontal: "center", vertical: "center" }, border: { top: {style:"thin"}, bottom: {style:"thin"}, left: {style:"thin"}, right: {style:"thin"} } };
        
        const sRed = { font: { color: { rgb: "C0392B" }, bold: true }, border: { top: {style:"thin"}, bottom: {style:"thin"}, left: {style:"thin"}, right: {style:"thin"} } };
        const sBlack = { font: { color: { rgb: "000000" }, bold: true }, border: { top: {style:"thin"}, bottom: {style:"thin"}, left: {style:"thin"}, right: {style:"thin"} } };
        const sBlue = { font: { color: { rgb: "00008B" }, bold: true }, border: { top: {style:"thin"}, bottom: {style:"thin"}, left: {style:"thin"}, right: {style:"thin"} } };

        let aoa = []; aoa.push([document.getElementById('summaryTitle').innerText]); aoa.push([dateRangeText]);
        const percent = totals.rev > 0 ? (totals.th / totals.rev * 100).toFixed(1) : 0; aoa.push([`Doanh số: ${formatNumberFull(totals.rev)} - Thưởng: ${formatNumberFull(totals.th)} (${percent}%)`]); aoa.push([""]); 
        
        let headers;
        if(appMode === 'excel') { headers = ["STT", "Đại Lý", "Lô tô", "Lô cặp", "Lô 3/23", "Lô 2/27"]; if (hasXSKT) headers.push("XSKT"); }
        else { headers = ["STT", "Đại Lý", "Doanh Số", "Hoa hồng", "Thưởng"]; }
        if(appMode==='excel') headers.push("Doanh số", "Thưởng");
        if (hasDebt) headers.push("Nợ cũ"); headers.push("Tiền nộp"); aoa.push(headers);
        
        let rowMetaData = {}; // rowIndex -> type
        let currentRow = 5; // Dữ liệu bắt đầu từ dòng 5 (index 0-based)

        let stt = 1; const sortMode = DOM.sortSelect.value; let sortedGroups = [...groups.entries()];
        if (sortMode === 'revenue_desc') sortedGroups.sort((a, b) => b[1].gr - a[1].gr); else if (sortMode === 'revenue_asc') sortedGroups.sort((a, b) => a[1].gr - b[1].gr);
        
        // 1. VÒNG LẶP DỮ LIỆU
        sortedGroups.forEach(([gk, g]) => { 
            const isMulti = g.agents.length > 1;
            g.agents.forEach((ag, idx) => { 
                const d = ag.data; const final = (d.totRev - d.totCom - d.thuong) + d.debt; 
                let displayStt = (idx === 0) ? stt : ""; 
                let row;
                if(appMode === 'excel') { row = [ displayStt, ag.code, d.lotoDB, d.lotoCap, d.c323, d.c227 ]; if (hasXSKT) row.push(d.xskt); row.push(d.totRev, d.thuong); }
                else { row = [ displayStt, ag.code, d.totRev, d.totCom, d.thuong ]; }
                if (hasDebt) row.push(d.debt); row.push(final); 
                
                aoa.push(row); 
                rowMetaData[currentRow] = isMulti ? 'group_member' : 'single';
                currentRow++;
            }); 
            
            if (isMulti) { 
                let rowGroup = ["", ""]; 
                if(appMode==='excel') { rowGroup.push("","","",""); if(hasXSKT) rowGroup.push(""); rowGroup.push(g.gr, g.gth); }
                else { rowGroup.push(g.gr, "", g.gth); }
                if (hasDebt) { let sumDebt = 0; g.agents.forEach(ag => sumDebt+=ag.data.debt); rowGroup.push(sumDebt); } rowGroup.push(g.gt); 
                
                aoa.push(rowGroup); 
                rowMetaData[currentRow] = 'group_total';
                currentRow++;
            } 
            stt++; 
        }); 

        // 2. THÊM DÒNG TỔNG CỘNG (GRAND TOTAL) - [FIXED]
        let rowTotal = ["TỔNG", ""]; // Cột 0 là chữ TỔNG, Cột 1 để trống để merge
        if(appMode === 'excel') { 
            rowTotal.push(totals.lotoDB, totals.lotoCap, totals.c323, totals.c227); 
            if(hasXSKT) rowTotal.push(totals.xskt); 
            rowTotal.push(totals.rev, totals.th); 
        } else { 
            rowTotal.push(totals.rev, totals.com, totals.th); 
        }
        if (hasDebt) rowTotal.push(totals.debt); 
        rowTotal.push(totals.final);
        
        aoa.push(rowTotal);
        let totalRowIndex = currentRow; // Lưu chỉ số dòng tổng để style
        // currentRow++; 

        const wb = XLSX.utils.book_new(); const ws = XLSX.utils.aoa_to_sheet(aoa);
        const range = XLSX.utils.decode_range(ws['!ref']); const lastColIdx = range.e.c;
        if(!ws['!merges']) ws['!merges'] = [];
        
        // Merge header
        ws['!merges'].push({ s: { r: 0, c: 0 }, e: { r: 0, c: lastColIdx } }); 
        ws['!merges'].push({ s: { r: 1, c: 0 }, e: { r: 1, c: lastColIdx } }); 
        ws['!merges'].push({ s: { r: 2, c: 0 }, e: { r: 2, c: lastColIdx } }); 
        ws['!merges'].push({ s: { r: 4, c: 0 }, e: { r: 4, c: 1 } }); // Merge STT & ĐL ở header nếu cần (tùy chọn)
        
        // Merge dòng TỔNG ở cuối [FIXED]
        ws['!merges'].push({ s: { r: totalRowIndex, c: 0 }, e: { r: totalRowIndex, c: 1 } });

        // 3. ÁP DỤNG STYLE
        for (let R = range.s.r; R <= range.e.r; ++R) { 
            for (let C = range.s.c; C <= range.e.c; ++C) { 
                const cellRef = XLSX.utils.encode_cell({ r: R, c: C }); 
                if (!ws[cellRef]) continue; 
                
                if (R <= 2) { 
                    ws[cellRef].s = { font: { bold: true, sz: 14, color: {rgb: R===2?"E74C3C":"E67E22"} }, alignment: { horizontal: "center" } }; if(R===1) ws[cellRef].s.font = { italic: true, sz: 11 }; 
                } 
                else if (R === 4) { 
                    ws[cellRef].s = sHeader; 
                } 
                else if (R === totalRowIndex) { 
                    // Style cho dòng TỔNG cuối cùng [FIXED]
                    ws[cellRef].s = sTotal;
                    // Căn giữa chữ "TỔNG"
                    if (C === 0) ws[cellRef].s = { ...sTotal, alignment: { horizontal: "center", vertical: "center" } };
                    
                    // Màu số tiền cuối cùng (Dương đỏ / Âm xanh)
                    if (C === lastColIdx) {
                        if (ws[cellRef].v < 0) ws[cellRef].s = { ...sTotal, font: { bold: true, color: { rgb: "00008B" } } }; // Blue
                        else ws[cellRef].s = { ...sTotal, font: { bold: true, color: { rgb: "C0392B" } } }; // Red
                    }
                }
                else { 
                    const meta = rowMetaData[R];
                    const isLastCol = (C === lastColIdx);
                    
                    if (meta === 'group_total') { 
                        ws[cellRef].s = sGroup; 
                        if(isLastCol) { 
                             if (ws[cellRef].v < 0) ws[cellRef].s = sBlue;
                             else ws[cellRef].s = { ...sGroup, font: { bold: true, color: { rgb: "C0392B" } } }; 
                        } 
                    } 
                    else { // Single or Group Member
                        ws[cellRef].s = sNormal; 
                        
                        // [FIXED] Style cho cột STT (0) và ĐL (1): In đậm + Căn giữa
                        if (C === 0 || C === 1) {
                            ws[cellRef].s = sCenterBold;
                        }

                        if (isLastCol) {
                            if (meta === 'group_member') {
                                ws[cellRef].s = sBlack;
                            } else {
                                // Đại lý lẻ (single): Đỏ (dương) / Xanh đậm (âm)
                                if (ws[cellRef].v < 0) ws[cellRef].s = sBlue; 
                                else ws[cellRef].s = sRed; // Positive -> Red
                            }
                        } 
                    } 
                } 
                
                if (typeof ws[cellRef].v === 'number') { ws[cellRef].z = '#,##0'; } 
            } 
        }
        ws['!cols'] = [ { wch: 5 }, { wch: 15 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 12 }, { wch: 12 }, { wch: 12 }, { wch: 15 } ];
        
        XLSX.utils.book_append_sheet(wb, ws, "Báo Cáo"); 
        
        // --- XỬ LÝ TÊN FILE ---
        const unitName = removeVietnameseTones(DOM.regionSelect.options[DOM.regionSelect.selectedIndex].text).replace(/\s+/g, '_');
        let safeDate = removeVietnameseTones(dateRangeText).replace(/[\\/]/g, "-").replace(/[:]/g, "").replace(/\s+/g, "_");
        if(safeDate === "Chua_co_ngay") safeDate = "Moi_nhat";
        const fileName = `bao_cao_${unitName}_${safeDate}.xlsx`;
        
        XLSX.writeFile(wb, fileName);
    }

    DOM.btnGoToSummary.onclick=()=>{ showSection('summarySection'); }; DOM.btnGoToSell.onclick=()=>{ showSection('sellSection'); DOM.sellQuantityInput.focus(); }; DOM.btnBackFromSummary.onclick=()=>{ showSection('reportSection'); DOM.soSoInput.focus(); }; DOM.btnBackFromSell.onclick=()=>{ showSection('reportSection'); DOM.soSoInput.focus(); };
    DOM.excelFile1.onchange=(e)=>handleFile(e.target.files[0],1,'excelFile1'); DOM.excelFile2.onchange=(e)=>handleFile(e.target.files[0],2,'excelFile2');
    
    // --- UPDATE SỰ KIỆN CLICK CHO 2 NÚT MỚI ---
    if(DOM.btnCreateDailyReport) DOM.btnCreateDailyReport.onclick = () => createReport('detailed');
    if(DOM.btnCreateSummaryReport) DOM.btnCreateSummaryReport.onclick = () => createReport('summary');
    
    DOM.soSoInput.onkeydown = (e) => { 
        if (e.key === 'Enter') { 
            e.preventDefault(); 
            e.target.blur(); 
            // Nếu nút báo cáo ngày ẩn (chế độ excel) -> chạy summary, ngược lại chạy detailed
            if (DOM.btnCreateDailyReport && DOM.btnCreateDailyReport.style.display === 'none') {
                createReport('summary');
            } else {
                createReport('detailed'); 
            }
        } 
    };

    if(DOM.btnShowManualReport_inline) { DOM.btnShowManualReport_inline.onclick = () => { DOM.manualInputModal.style.display = 'block'; DOM.manualAgentName.focus(); }; }
    DOM.btnShowManualReport.onclick=()=>{DOM.manualInputModal.style.display='block';DOM.manualAgentName.focus()}; DOM.closeModalBtn.onclick=()=>{DOM.manualInputModal.style.display='none'};
    DOM.btnCreateManualReport.onclick=()=>{ const name=DOM.manualAgentName.value; if(!name) return; const t={ lotoDB:(parseFloat(DOM.manualLotoDB.dataset.value)||0)*1000, loCap:(parseFloat(DOM.manualLoCap.dataset.value)||0)*1000, c227:(parseFloat(DOM.manualC227.dataset.value)||0)*1000, c323:(parseFloat(DOM.manualC323.dataset.value)||0)*1000, thuong:(parseFloat(DOM.manualThuong.dataset.value)||0)*1000, xskt:0 }; const c={ldb:t.lotoDB*.08, lc:t.loCap*.1, c2:t.c227*.1, c3:t.c323*.1, x:0}; resetState('manual'); state.accountCodes.add(name); state.baseTongDT=t.lotoDB+t.loCap+t.c227+t.c323; state.baseTongHH=c.ldb+c.lc+c.c2+c.c3; state.baseThuong=t.thuong; DOM.homePage.style.display='none'; DOM.mainPage.style.display='block'; showSection('reportSection'); renderDetailedReport({t,c},DOM.manualDateRange.value); updateDetailedView(); DOM.manualInputModal.style.display='none'; DOM.actionButtonContainer.style.display='grid'; DOM.actionChoices.style.display='none'; DOM.toggleChoicesBtn.textContent='Thêm tiền Thưởng/Nợ cũ'; DOM.toggleChoicesBtn.className='secondary'; };
    document.querySelectorAll('.manual-currency').forEach(i=>i.oninput=(e)=>{const v=e.target.value.replace(/\D/g,'');e.target.dataset.value=v;e.target.value=v?parseInt(v).toLocaleString('vi-VN'):'';});
    DOM.btnExportExcel.onclick = exportSummaryToExcel; DOM.btnPrintReport.onclick = printReport; DOM.btnCopySummary.onclick=()=>copyElementAsImage('#summarySection','Đã copy ảnh',DOM.btnCopySummary); DOM.btnCopyReport.onclick=()=>copyElementAsImage('#reportContainer','Đã copy ảnh',DOM.btnCopyReport);
    DOM.debtModeRadios.forEach(r=>r.onchange=(e)=>{ DOM.debtManualContainer.style.display=e.target.value==='manual'?'block':'none'; DOM.debtExcelContainer.style.display=e.target.value==='excel'?'block':'none'; DOM.btnApplyDebt.style.display=e.target.value==='manual'?'block':'none'; debtData.clear(); DOM.debtStatusMsg.style.display='none'; generateSummaryView();});
    DOM.debtFile.onchange=(e)=>handleDebtFile(e.target.files[0]); DOM.btnClearDebt.onclick=()=>{debtData.clear(); DOM.debtInput.value=''; DOM.debtFile.value=''; DOM.debtManualContainer.style.display='none';DOM.debtExcelContainer.style.display='none';DOM.debtStatusMsg.style.display='none';DOM.debtModeRadios.forEach(r=>r.checked=false); DOM.btnApplyDebt.style.display='none'; generateSummaryView();}; DOM.btnApplyDebt.onclick=parseDebtData;
    DOM.btnSaveRegions.onclick=()=>{ if(isRegionLocked) { DOM.passwordModal.style.display='block'; DOM.passwordInput.focus(); } else parseRegions(); };
    DOM.btnSubmitPassword.onclick=()=>{ if(DOM.passwordInput.value==='admin'){ DOM.passwordModal.style.display='none'; isRegionLocked=false; DOM.regionLockContainer.classList.remove('locked'); DOM.btnSaveRegions.textContent='Lưu Định nghĩa'; DOM.btnSaveRegions.classList.remove('warning'); DOM.btnSaveRegions.classList.add('success'); } else showNotification('Sai mật khẩu','error'); };
    DOM.closePasswordModalBtn.onclick=()=>{DOM.passwordModal.style.display='none'};
    function parseRegions(){ const txt=DOM.regionDefinitionsInput.value; regionDefinitions={}; const lines=txt.split('\n'); DOM.regionSelect.innerHTML='<option value="all">--- Tất cả ---</option>'; lines.forEach(l=>{ const idx=l.indexOf(':'); if(idx>0){ const n=l.slice(0,idx).trim(), r=l.slice(idx+1).trim(); const rules=[]; r.split('+').forEach(p=>{ if(p.includes('_')){const[mi,ma]=p.split('_'); rules.push({min:+mi,max:+ma});} else rules.push({min:+p,max:+p}); }); if(rules.length){ regionDefinitions[n]=rules; const o=document.createElement('option'); o.value=n; o.text=n; DOM.regionSelect.add(o); } }}); localStorage.setItem('zengRegions',txt); isRegionLocked=true; DOM.regionLockContainer.classList.add('locked'); DOM.btnSaveRegions.textContent='Mở khóa'; DOM.btnSaveRegions.classList.replace('success','warning'); }
    async function fetchDefaultRegions() { try { const response = await fetch(`${REGION_DATA_URL}?t=${Date.now()}`); if (!response.ok) throw new Error('Không thể tải file'); const text = await response.text(); DOM.regionDefinitionsInput.value = text; parseRegions(); } catch (error) { console.warn('Lỗi tải cấu hình online:', error); const saved = localStorage.getItem('zengRegions'); if(saved){ DOM.regionDefinitionsInput.value = saved; parseRegions(); showNotification('Đang dùng cấu hình Offline', 'info'); } } }
    fetchDefaultRegions();
    DOM.regionSelect.onchange=generateSummaryView; DOM.sortSelect.onchange=generateSummaryView; DOM.minRevenueInput.onblur=generateSummaryView; DOM.maxRevenueInput.onblur=generateSummaryView;
    function updateFilterStatus() { const isDefault = DOM.sortSelect.value === 'default' && !DOM.minRevenueInput.dataset.value && !DOM.maxRevenueInput.dataset.value && DOM.regionSelect.value === 'all'; const btn = DOM.btnClearRevenueFilter; btn.className = isDefault ? 'secondary' : 'danger'; }
    DOM.regionSelect.addEventListener('change', updateFilterStatus); DOM.sortSelect.addEventListener('change', updateFilterStatus); DOM.minRevenueInput.addEventListener('blur', updateFilterStatus); DOM.maxRevenueInput.addEventListener('blur', updateFilterStatus);
    DOM.btnClearRevenueFilter.onclick=()=>{ DOM.sortSelect.value='default'; DOM.minRevenueInput.value=''; DOM.minRevenueInput.dataset.value=''; DOM.maxRevenueInput.value=''; DOM.maxRevenueInput.dataset.value=''; DOM.regionSelect.value='all'; generateSummaryView(); updateFilterStatus(); };
    window.onscroll=()=>{ const s=window.scrollY; DOM.scrollToTopBtn.style.display=s>200?'block':'none'; DOM.scrollToBottomBtn.style.display=s<(document.body.scrollHeight-window.innerHeight-200)?'block':'none'; }; DOM.scrollToTopBtn.onclick=()=>window.scrollTo({top:0,behavior:'smooth'}); DOM.scrollToBottomBtn.onclick=()=>window.scrollTo({top:document.body.scrollHeight,behavior:'smooth'});
    DOM.sellTicketTypeRadios.forEach(r=>r.addEventListener('change', (e)=>{ const v=e.target.value; DOM.customTicketName.style.display=v==='custom-name'?'block':'none'; if(v==='xo-so'){ const pRad = document.querySelector('input[name="sellPrice"][value="10000"]'); if(pRad) pRad.checked = true; DOM.customTicketPrice.style.display = 'none'; const rRad = document.querySelector('input[name="sellRate"][value="10"]'); if(rRad) rRad.checked = true; DOM.customTicketRate.style.display = 'none'; DOM.sellQuantityInput.focus(); } else if(v==='ve-boc'){ const pRad = document.querySelector('input[name="sellPrice"][value="5000"]'); if(pRad) pRad.checked = true; DOM.customTicketPrice.style.display = 'none'; const rRad = document.querySelector('input[name="sellRate"][value="12"]'); if(rRad) rRad.checked = true; DOM.customTicketRate.style.display = 'none'; DOM.sellQuantityInput.focus(); } else if(v==='custom-name') { DOM.customTicketName.focus(); } }));
    DOM.sellPriceRadios.forEach(r=>r.addEventListener('change', (e)=>{ DOM.customTicketPrice.style.display=e.target.value==='custom-price'?'block':'none'; if(e.target.value==='custom-price') DOM.customTicketPrice.focus(); else DOM.sellQuantityInput.focus(); }));
    DOM.sellRateRadios.forEach(r=>r.addEventListener('change', (e)=>{ DOM.customTicketRate.style.display=e.target.value==='custom-rate'?'block':'none'; if(e.target.value==='custom-rate') DOM.customTicketRate.focus(); else DOM.sellQuantityInput.focus(); }));
    DOM.customTicketName.addEventListener('keydown', (e)=>{ if(e.key==='Enter'){ e.preventDefault(); if(DOM.customTicketPrice.style.display!=='none') DOM.customTicketPrice.focus(); else DOM.sellQuantityInput.focus(); } }); DOM.customTicketPrice.addEventListener('keydown', (e)=>{ if(e.key==='Enter'){ e.preventDefault(); DOM.sellQuantityInput.focus(); } }); DOM.customTicketRate.addEventListener('keydown', (e)=>{ if(e.key==='Enter'){ e.preventDefault(); DOM.sellQuantityInput.focus(); } }); DOM.customTicketPrice.oninput=(e)=>{e.target.value=e.target.value.replace(/\D/g,'').replace(/\B(?=(\d{3})+(?!\d))/g,".")}; DOM.sellQuantityInput.addEventListener('keydown', (e) => { if(e.key === 'Enter') { e.preventDefault(); addTicket(); } }); DOM.btnAddTicket.onclick = addTicket;
    manualFormOrder.forEach((id, index) => { const el = document.getElementById(id); if (el) { el.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); if (index < manualFormOrder.length - 1) { document.getElementById(manualFormOrder[index + 1]).focus(); } else { DOM.btnCreateManualReport.click(); } } }); } }); DOM.passwordInput.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); DOM.btnSubmitPassword.click(); } });
    DOM.btnClearManualReport.onclick = () => { const ids = ['manualAgentName', 'manualDateRange', 'manualLotoDB', 'manualLoCap', 'manualC227', 'manualC323', 'manualThuong']; ids.forEach(id => { const el = document.getElementById(id); if (el) { el.value = ''; if (el.dataset.value) el.dataset.value = ''; } }); DOM.manualAgentName.focus(); }; document.querySelectorAll('.simple-animated-logo').forEach(el => el.addEventListener('click', () => location.reload()));

    runLogoAnimation(); updateOnlineCounter();
})();