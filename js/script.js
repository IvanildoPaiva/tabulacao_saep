let buffMestre, dadosRaw, nomeMestre;

// Event Listeners
document.getElementById('file-mestre').addEventListener('change', (e) => load(e.target.files[0], 'mestre'));
document.getElementById('file-dados').addEventListener('change', (e) => load(e.target.files[0], 'dados'));

function load(f, type) {
    if(!f) return;
    let r = new FileReader();
    r.onload = (e) => {
        if(type === 'mestre') {
            buffMestre = e.target.result;
            nomeMestre = f.name;
            document.getElementById('box-mestre').classList.add('ok');
            document.getElementById('box-mestre').innerText = "✅ Mestre OK: " + f.name;
        } else {
            // Ler dados brutos com SheetJS
            let wb = XLSX.read(new Uint8Array(e.target.result), {type:'array'});
            let ws = wb.Sheets[wb.SheetNames[1] || wb.SheetNames[0]];
            dadosRaw = XLSX.utils.sheet_to_json(ws, {header:1, defval:''});
            document.getElementById('box-dados').classList.add('ok');
            document.getElementById('box-dados').innerText = "✅ Dados OK: " + f.name;
        }
    };
    r.readAsArrayBuffer(f);
}

async function processar() {
    if(!buffMestre || !dadosRaw) return alert("Por favor, carregue os dois arquivos.");
    
    let lg = document.getElementById('console'); 
    lg.style.display='block';
    lg.innerText = "⏳ Iniciando processamento completo...";

    try {
        // Carrega Mestre com XlsxPopulate (Preserva Gráficos)
        const workbook = await XlsxPopulate.fromDataAsync(buffMestre);
        const sheet = workbook.sheet("AVALIAÇÃO PO");
        if(!sheet) throw new Error("Aba 'AVALIAÇÃO PO' não encontrada!");

        // --- CONFIGURAÇÃO FIXA ---
        const C_ID=0, C_CAP=1, C_DIF=5, C_RESP=6, C_GAB=7;

        // 1. Limpeza
        lg.innerText += "\n🧹 Limpando dados...";
        for(let r=8; r<=47; r++) {
            sheet.row(r).cell(2).value(null);
            sheet.row(r).cell(3).value(null);
            for(let c=4; c<=43; c++) sheet.row(r).cell(c).value(null);
        }
        for(let c=4; c<=43; c++) {
            sheet.row(4).cell(c).value(null);
            sheet.row(7).cell(c).value(null);
        }

        // 2. Leitura
        lg.innerText += "\n🔄 Processando alunos e descrições...";
        
        let alunos = new Map(), mapQ = new Map();
        let arrCap = Array(40).fill(null), arrDif = Array(40).fill(null);
        // Novo Mapa para guardar as descrições (Ex: "C1" -> "Utilizar aplicações...")
        let mapDescricoes = new Map();
        
        let curAluno = null;

        for(let i=0; i<dadosRaw.length; i++) {
            let r = dadosRaw[i];
            let val = String(r[C_ID]||'').trim();

            if(val === 'Aluno') {
                if(i+1 < dadosRaw.length) {
                    let nm = String(dadosRaw[i+1][C_ID]||'').trim();
                    if(nm) {
                        curAluno = nm;
                        if(!alunos.has(nm)) alunos.set(nm, Array(40).fill(0));
                    }
                }
                continue;
            }

            if(val.startsWith('SAEP_') && curAluno) {
                if(!mapQ.has(val) && mapQ.size < 40) mapQ.set(val, mapQ.size);
                let idx = mapQ.get(val);
                
                if(idx !== undefined) {
                    let capFull = String(r[C_CAP]||'').trim(); // Ex: "C3 - Aplicar lógica..."
                    let dif = String(r[C_DIF]||'').trim().toUpperCase();
                    let resp = String(r[C_RESP]||'').trim().toUpperCase();
                    let gab = String(r[C_GAB]||'').trim().toUpperCase();

                    // --- LÓGICA DE EXTRAÇÃO ---
                    let capCode = ""; // Só o C3
                    
                    if(capFull.includes("-")) {
                        let parts = capFull.split("-");
                        capCode = parts[0].trim(); // "C3"
                        
                        // Guarda a descrição (o que vem depois do hífen)
                        let capDesc = parts.slice(1).join("-").trim(); // "Aplicar lógica..."
                        if(capCode && capDesc && !mapDescricoes.has(capCode)) {
                            mapDescricoes.set(capCode, capDesc);
                        }
                    } else {
                        capCode = capFull.split(" ")[0].trim();
                    }

                    if(!arrCap[idx]) arrCap[idx] = capCode;

                    if(!arrDif[idx]) {
                         if(dif.includes("FACIL") || dif.includes("FÁCIL")) arrDif[idx] = dif.includes("MUITO") ? "MF" : "F";
                        else if(dif.includes("MEDIO") || dif.includes("MÉDIO")) arrDif[idx] = "M";
                        else if(dif.includes("DIFICIL") || dif.includes("DIFÍCIL")) arrDif[idx] = dif.includes("MUITO") ? "MD" : "D";
                    }
                    
                    alunos.get(curAluno)[idx] = (resp===gab && resp!=='') ? 1 : 0;
                }
            }
        }

        // 3. Escrever na AVALIAÇÃO PO
        lg.innerText += "\n📝 Preenchendo notas...";
        arrCap.forEach((v,k) => { if(v) sheet.row(4).cell(4+k).value(v); });
        arrDif.forEach((v,k) => { if(v) sheet.row(7).cell(4+k).value(v); });

        let lin = 8, count = 1;
        for(let [nome, notas] of alunos) {
            if(lin > 47) break;
            sheet.row(lin).cell(2).value(count);
            sheet.row(lin).cell(3).value(nome);
            notas.forEach((n,k) => sheet.row(lin).cell(4+k).value(n));
            lin++; count++;
        }

        // 4. Preencher Descrições na DIAGNÓSTICO PO (NOVA FUNCIONALIDADE)
        const sheetDiag = workbook.sheet("DIAGNÓSTICO PO");
        if(sheetDiag) {
            lg.innerText += "\n📋 Preenchendo aba de Diagnóstico...";
            
            // Varre a Coluna B (onde ficam C1, C2...) e preenche a C (Descrição)
            // Vamos assumir que a tabela está entre a linha 4 e 30
            for(let r=4; r<=30; r++) {
                let cellKey = sheetDiag.row(r).cell(2).value(); // Lê Coluna B (2)
                
                if(cellKey && typeof cellKey === 'string') {
                    cellKey = cellKey.trim(); // Garante que é "C1" limpo
                    
                    if(mapDescricoes.has(cellKey)) {
                        // Escreve a descrição na Coluna C (3)
                        sheetDiag.row(r).cell(3).value(mapDescricoes.get(cellKey));
                    }
                }
            }
        }

        // 5. Salvar
        const blob = await workbook.outputAsync();
        saveAs(blob, nomeMestre.replace('.xlsx', '_FINAL.xlsx'));
        
        lg.innerText += "\n✅ Sucesso! Gráficos e Descrições atualizados.";

    } catch(e) {
        lg.innerText += "\n❌ Erro Fatal: " + e.message;
        console.error(e);
    }
}
