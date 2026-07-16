// Marguerite A Gemi Makine Hiyerarşisi (Varsayılan / Demo Veri Seti - 20 Kritik Parça)
const defaultVesselData = [
    // Seviye 1: Ana Makine (MAN B&W 6S50ME-C)
    { id: "ME-100", name: "Silindir Gömleği (Cylinder Liner)", parent: "Ana Makine", hours: 14200, maxHours: 16000, desc: "Kritik sürtünme yüzeyi. Aşınma toleransı sınırda." },
    { id: "ME-200", name: "Piston Kafası (Piston Crown)", parent: "Ana Makine", hours: 9800, maxHours: 12000, desc: "Yüksek sıcaklık dayanımı kontrol edilmeli." },
    { id: "ME-300", name: "Egzoz Valfi (Exhaust Valve Spindle)", parent: "Ana Makine", hours: 5100, maxHours: 6000, desc: "Kurum birikmesi ve sızdırmazlık testi önerilir." },
    { id: "ME-400", name: "Yakıt Enjektörü (Fuel Injector Nozzle)", parent: "Ana Makine", hours: 1950, maxHours: 2000, desc: "Püskürtme kalitesi bozulmuş, acil değişim gerekli!" },
    { id: "ME-500", name: "Krank Mili Yatağı (Main Bearing)", parent: "Ana Makine", hours: 22000, maxHours: 30000, desc: "Yağ analizi temiz, mekanik aşınma normal düzeyde." },

    // Seviye 2: Yardımcı Jeneratör (Auxiliary Engine No:1)
    { id: "AE-100", name: "Turboşarj Kartuşu (Turbocharger Rotor)", parent: "Yardımcı Jeneratör", hours: 7800, maxHours: 8000, desc: "Rulman titreşim seviyesi yüksek." },
    { id: "AE-200", name: "Enjektör Pompası (Fuel Injection Pump)", parent: "Yardımcı Jeneratör", hours: 4100, maxHours: 10000, desc: "Performans stabil." },
    { id: "AE-300", name: "Alternatör Rulmanı (Alternator Bearing)", parent: "Yardımcı Jeneratör", hours: 15500, maxHours: 20000, desc: "Sıcaklık değerleri normal." },

    // Seviye 3: Seperatörler ve Yakıt Arıtma (Purifiers)
    { id: "PR-100", name: "LO Seperatör Tamburu (Lub Oil Purifier Bowl)", parent: "Arıtma Sistemleri", hours: 3200, maxHours: 4000, desc: "Balans ayarı gerekebilir." },
    { id: "PR-200", name: "HFO Seperatör Dişli Kutusu (Fuel Purifier Gearbox)", parent: "Arıtma Sistemleri", hours: 9100, maxHours: 12000, desc: "Yağ değişimi yapıldı, dişliler sağlam." },

    // Seviye 4: Balast ve Pompa Sistemleri
    { id: "PM-100", name: "Balast Pompası Fanı (Ballast Pump Impeller)", parent: "Pompa Sistemleri", hours: 2100, maxHours: 5000, desc: "Kavitasyon izleri mevcut ancak çalışır durumda." },
    { id: "PM-200", name: "Ana Deniz Suyu Pompası Şaftı (SW Pump Shaft)", parent: "Pompa Sistemleri", hours: 8200, maxHours: 10000, desc: "Sızdırmazlık salmastraları kontrol edilmeli." }
];

let activeData = [];

document.addEventListener("DOMContentLoaded", () => {
    // Event Listeners
    document.getElementById("excelFile").addEventListener("change", handleExcelUpload);
    document.getElementById("demoBtn").addEventListener("click", loadDemoData);
});

// Demo Veriyi Yükleme Fonksiyonu
function loadDemoData() {
    activeData = JSON.parse(JSON.stringify(defaultVesselData));
    processAndRenderData();
}

// Excel Dosyasını Okuma
function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function(evt) {
        const data = evt.target.result;
        const workbook = XLSX.read(data, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // Excel verilerini JSON'a çevir
        const excelRows = XLSX.utils.sheet_to_json(worksheet);
        
        if (excelRows.length > 0) {
            // Excel kolon isimlerini bizim şemaya map'le
            activeData = excelRows.map((row, index) => ({
                id: row["Parça Kodu"] || row["ID"] || `PART-${100 + index}`,
                name: row["Parça Adı"] || row["Name"] || "Bilinmeyen Parça",
                parent: row["Sistem"] || row["Parent"] || "Diğer",
                hours: parseInt(row["Çalışma Saati"] || row["Hours"]) || 0,
                maxHours: parseInt(row["Ömür Sınırı"] || row["MaxHours"]) || 10000,
                desc: row["Açıklama"] || row["Description"] || "AI analizi bekleniyor..."
            }));
            processAndRenderData();
        } else {
            alert("Boş veya uyumsuz bir Excel dosyası!");
        }
    };
    reader.readAsBinaryString(file);
}

// Verileri İşleme ve Görselleştirme
function processAndRenderData() {
    calculateStats();
    renderTreeView();
    renderTable();
}

// İstatistikleri Hesapla (Üst Widget Kartları)
function calculateStats() {
    const total = activeData.length;
    let criticalCount = 0;
    let totalLifetime = 0;

    activeData.forEach(item => {
        const usageRatio = item.hours / item.maxHours;
        // AI Simülasyonu: %85 kullanımın üzerini "Kritik" kabul et
        if (usageRatio >= 0.85) {
            criticalCount++;
        }
        totalLifetime += item.maxHours;
    });

    const avgLife = total > 0 ? Math.round(totalLifetime / total) : 0;

    document.getElementById("totalParts").textContent = total;
    document.getElementById("criticalParts").textContent = criticalCount;
    document.getElementById("avgLifetime").textContent = `${avgLife.toLocaleString()} Saat`;
}

// Soy Ağacı (Tree View) Oluşturma
function renderTreeView() {
    const treeContainer = document.getElementById("treeContainer");
    treeContainer.innerHTML = "";

    // Benzersiz ana sistemleri grupla
    const systems = [...new Set(activeData.map(item => item.parent))];

    systems.forEach(system => {
        // Ana Düğüm (Parent Node)
        const parentDiv = document.createElement("div");
        parentDiv.className = "mb-4";
        parentDiv.innerHTML = `
            <div class="flex items-center gap-2 p-2 bg-slate-900 rounded-lg border border-slate-800 cursor-pointer hover:bg-slate-800 transition">
                <i class="fa-solid fa-folder-open text-blue-400"></i>
                <span class="font-bold text-sm text-blue-300">${system}</span>
            </div>
            <div class="space-y-1 mt-1 ml-2 border-l border-slate-800" id="children-of-${system.replace(/\s+/g, '-')}">
            </div>
        `;
        treeContainer.appendChild(parentDiv);

        // Alt Düğümler (Child Nodes)
        const childrenContainer = document.getElementById(`children-of-${system.replace(/\s+/g, '-')}`);
        const children = activeData.filter(item => item.parent === system);

        children.forEach(child => {
            const childDiv = document.createElement("div");
            childDiv.className = "tree-node-child py-1 text-xs text-slate-300 flex items-center justify-between hover:text-white transition cursor-pointer";
            
            const remainingPct = Math.max(0, 100 - Math.round((child.hours / child.maxHours) * 100));
            const statusColor = remainingPct < 15 ? 'text-red-500' : (remainingPct < 40 ? 'text-yellow-500' : 'text-emerald-500');

            childDiv.innerHTML = `
                <div class="flex items-center gap-2" onclick="openAiAnalysis('${child.id}')">
                    <i class="fa-solid fa-gear text-slate-500"></i>
                    <span>${child.name}</span>
                </div>
                <span class="text-[10px] font-semibold ${statusColor} px-2 py-0.5 bg-slate-900 rounded">${remainingPct}% Ömür</span>
            `;
            childrenContainer.appendChild(childDiv);
        });
    });
}

// Tabloyu Doldurma
function renderTable() {
    const tbody = document.getElementById("partsTableBody");
    tbody.innerHTML = "";

    activeData.forEach(item => {
        const remainingPct = Math.max(0, 100 - Math.round((item.hours / item.maxHours) * 100));
        
        let statusBadge = "";
        let pctColor = "";

        if (remainingPct <= 15) {
            statusBadge = `<span class="px-2 py-1 bg-red-950 text-red-400 border border-red-800 rounded-full text-[10px] font-bold">ACİL DEĞİŞİM</span>`;
            pctColor = "text-red-500 font-bold";
        } else if (remainingPct <= 40) {
            statusBadge = `<span class="px-2 py-1 bg-yellow-950 text-yellow-400 border border-yellow-800 rounded-full text-[10px] font-bold">BAKIM PLANLA</span>`;
            pctColor = "text-yellow-500";
        } else {
            statusBadge = `<span class="px-2 py-1 bg-emerald-950 text-emerald-400 border border-emerald-800 rounded-full text-[10px] font-bold">DURUM İYİ</span>`;
            pctColor = "text-emerald-400";
        }

        const tr = document.createElement("tr");
        tr.className = "hover:bg-slate-900/50 cursor-pointer transition";
        tr.onclick = () => openAiAnalysis(item.id);

        tr.innerHTML = `
            <td class="p-4 font-mono text-xs text-blue-400">${item.id}</td>
            <td class="p-4 font-semibold">${item.name}</td>
            <td class="p-4 text-slate-400 text-xs">${item.parent}</td>
            <td class="p-4">${item.hours.toLocaleString()} hrs</td>
            <td class="p-4 text-slate-500">${item.maxHours.toLocaleString()} hrs</td>
            <td class="p-4 ${pctColor}">${remainingPct}%</td>
            <td class="p-4">${statusBadge}</td>
        `;
        tbody.appendChild(tr);
    });
}

// Yapay Zeka Detay Analizi (Modal Açılışı)
function openAiAnalysis(partId) {
    const part = activeData.find(p => p.id === partId);
    if (!part) return;

    const remainingPct = Math.max(0, 100 - Math.round((part.hours / part.maxHours) * 100));
    const modal = document.getElementById("aiModal");
    const modalContent = document.getElementById("modalContent");

    let aiRecommendation = "";
    if (remainingPct <= 15) {
        aiRecommendation = `<p class="text-sm text-red-300 bg-red-950/40 p-3 rounded-lg border border-red-900/50"><i class="fa-solid fa-triangle-exclamation mr-1"></i> <strong>Kestirimci Bakım Algoritması Uyarısı:</strong> Parça ömrünün kritik sınırına gelinmiştir. Bir sonraki limanda parça tedariği yapılması ve aşınma ölçümlerinin derhal sisteme işlenmesi önerilir.</p>`;
    } else {
        aiRecommendation = `<p class="text-sm text-emerald-300 bg-emerald-950/40 p-3 rounded-lg border border-emerald-900/50"><i class="fa-solid fa-circle-check mr-1"></i> <strong>AI Tahmini:</strong> Mevcut çalışma parametreleri ve sıcaklık eğilimleri stabil. Yaklaşık ${(part.maxHours - part.hours)} saat daha güvenli operasyon öngörülmektedir.</p>`;
    }

    modalContent.innerHTML = `
        <div class="grid grid-cols-2 gap-4">
            <div class="bg-slate-900 p-3 rounded-lg border border-slate-800">
                <span class="text-xs text-slate-400">Sistem</span>
                <p class="font-semibold text-sm">${part.parent}</p>
            </div>
            <div class="bg-slate-900 p-3 rounded-lg border border-slate-800">
                <span class="text-xs text-slate-400">Parça Kodu / ID</span>
                <p class="font-semibold text-sm font-mono text-blue-400">${part.id}</p>
            </div>
        </div>
        <div class="bg-slate-900 p-4 rounded-lg border border-slate-800 space-y-2">
            <h4 class="font-bold text-sm text-white">${part.name}</h4>
            <p class="text-xs text-slate-400">${part.desc}</p>
        </div>
        <div class="space-y-1">
            <div class="flex justify-between text-xs text-slate-400">
                <span>Kalan Kullanım Ömrü Oranı</span>
                <span class="font-bold">${remainingPct}%</span>
            </div>
            <div class="w-full bg-slate-800 h-2.5 rounded-full overflow-hidden">
                <div class="h-full bg-gradient-to-r ${remainingPct < 15 ? 'from-red-600 to-red-400' : 'from-blue-600 to-emerald-400'}" style="width: ${remainingPct}%"></div>
            </div>
        </div>
        ${aiRecommendation}
    `;

    modal.classList.remove("hidden");
    modal.classList.add("flex");
}

function closeModal() {
    const modal = document.getElementById("aiModal");
    modal.classList.add("hidden");
    modal.classList.remove("flex");
}