const $ = (id) => document.getElementById(id);

// Aturan tetap (ala SIAP): A ≥85, AB ≥80, B ≥75, BC ≥70, C ≥60, D ≥40, E <40
const LADDER = [
	{ label: "A", min: 85 },
	{ label: "AB", min: 80 },
	{ label: "B", min: 75 },
	{ label: "BC", min: 70 },
	{ label: "C", min: 60 },
	{ label: "D", min: 40 },
	{ label: "E", min: -1 },
];

let fullAoa = [];
let headerRowIdx = 0;
let rawHeader = [];
let rawRows = [];
let rowMap = [];
let fileName = "";
let components = []; // [{id,label,col,weight}]
let cNama = -1;
let cNim = -1;
let students = []; // [{nim,name,scores:{compId:val}}]
let graphMode = "huruf"; // 'huruf' | 'angka'
let compSeq = 0;

function newCompId() {
	compSeq += 1;
	return "a" + Date.now().toString(36) + "_" + compSeq;
}
function escapeHtml(s) {
	return String(s ?? "").replace(
		/[&<>"']/g,
		(m) =>
			({
				"&": "&amp;",
				"<": "&lt;",
				">": "&gt;",
				'"': "&quot;",
				"'": "&#39;",
			})[m],
	);
}
function parseScore(v) {
	if (v === "" || v == null) return null;
	const n = parseFloat(String(v).replace(",", "."));
	return isNaN(n) ? null : n;
}
function finalOf(s) {
	let total = 0;
	for (const c of components) {
		const v = s.scores ? s.scores[c.id] : null;
		if (v == null) return null;
		total += v * (c.weight / 100);
	}
	return Math.round(total * 100) / 100;
}
function gradeOf(v) {
	if (v == null) return "-";
	for (const g of LADDER) {
		if (v >= g.min) return g.label;
	}
	return LADDER[LADDER.length - 1].label;
}
function rentang(i) {
	if (i === 0) return ">= " + LADDER[i].min;
	if (i === LADDER.length - 1) return "< " + LADDER[i - 1].min;
	return LADDER[i].min + " - < " + LADDER[i - 1].min;
}
function avgOf(values) {
	const a = values.filter((v) => v != null);
	if (!a.length) return null;
	return a.reduce((x, y) => x + y, 0) / a.length;
}
const fmt = (v, d = 2) => (v == null ? "-" : Number(v).toFixed(d));

// ---------- Upload (sama seperti tool Penilaian) ----------
const dropZone = $("dropZone"),
	fileInput = $("fileInput");
dropZone.onclick = () => fileInput.click();
dropZone.ondragover = (e) => {
	e.preventDefault();
	dropZone.style.background = "#e0f7ff";
};
dropZone.ondragleave = () => (dropZone.style.background = "#f8fdff");
dropZone.ondrop = (e) => {
	e.preventDefault();
	dropZone.style.background = "#f8fdff";
	if (e.dataTransfer.files[0]) readFile(e.dataTransfer.files[0]);
};
fileInput.onchange = (e) => {
	if (e.target.files[0]) readFile(e.target.files[0]);
};

function readFile(file) {
	fileName = file.name;
	const reader = new FileReader();
	reader.onload = (e) => {
		try {
			const wb = XLSX.read(e.target.result, { type: "array" });
			const ws = wb.Sheets[wb.SheetNames[0]];
			const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
			if (aoa.length < 2) {
				alert("Excel kosong / tidak ada data.");
				return;
			}
			if (autoApply(aoa)) {
				$("fileInfo").textContent =
					"File: " + fileName + " | " + students.length + " mahasiswa.";
			}
		} catch (err) {
			alert("Gagal membaca file: " + err.message);
		}
	};
	reader.readAsArrayBuffer(file);
}

function detectHeaderRow(aoa) {
	const limit = Math.min(aoa.length, 20);
	let best = 0,
		bestScore = -1;
	for (let r = 0; r < limit; r++) {
		const cells = aoa[r].map((c) =>
			String(c ?? "")
				.toLowerCase()
				.trim(),
		);
		const joined = cells.join(" | ");
		let score = 0;
		const hasNim = cells.some((c) => c === "nim" || c.includes("nim"));
		const hasNama = cells.some(
			(c) =>
				c.includes("nama mahasiswa") ||
				c === "nama" ||
				(c.includes("nama") && c.length < 30),
		);
		if (hasNim) score += 10;
		if (hasNama) score += 10;
		if (hasNim && hasNama) score += 20;
		if (joined.includes("mata kuliah") && !hasNim) score -= 20;
		if (joined.includes("tahun ajaran") && !hasNim) score -= 20;
		if (cells.filter((c) => c !== "").length < 2) score -= 10;
		if (score > bestScore) {
			bestScore = score;
			best = r;
		}
	}
	if (bestScore < 10) return 0;
	return best;
}

function guessIndex(patterns) {
	for (let i = 0; i < rawHeader.length; i++) {
		const h = rawHeader[i].toLowerCase();
		if (patterns.some((p) => h.includes(p))) return i;
	}
	return -1;
}
function cleanName(header) {
	let s = String(header ?? "")
		.split("(")[0]
		.trim();
	return s.replace(/^nilai\s+/i, "").trim();
}
function parseWeight(header) {
	const m = String(header ?? "").match(/\((\d+(?:[.,]\d+)?)\s*%\)\s*$/);
	return m ? parseFloat(m[1].replace(",", ".")) : 0;
}
function isScoreCol(h) {
	const hl = String(h ?? "").toLowerCase();
	return (
		hl.startsWith("nilai") &&
		!["akhir", "huruf", "bobot"].some((k) => hl.includes(k))
	);
}

// Otomatis: deteksi header → NIM/Nama → komponen + bobot → hitung
function autoApply(aoa) {
	fullAoa = aoa;
	headerRowIdx = detectHeaderRow(aoa);
	rawHeader = (fullAoa[headerRowIdx] || []).map((h) => String(h ?? "").trim());
	rawRows = [];
	rowMap = [];
	fullAoa.slice(headerRowIdx + 1).forEach((r, k) => {
		if (r.some((c) => String(c ?? "").trim() !== "")) {
			rowMap.push(headerRowIdx + 1 + k);
			rawRows.push(r);
		}
	});

	cNim = guessIndex(["nim"]);
	cNama = guessIndex(["nama mahasiswa", "nama", "name"]);
	if ((cNim < 0 || cNama < 0) && rawHeader.length >= 2) {
		if (cNim < 0) cNim = 0;
		if (cNama < 0) cNama = 1;
	}
	if (cNim < 0 || cNama < 0 || cNim === cNama) {
		alert("Kolom NIM / Nama tidak dikenali di file ini.");
		return false;
	}

	components = [];
	rawHeader.forEach((h, i) => {
		if (isScoreCol(h)) {
			components.push({
				id: newCompId(),
				label: cleanName(h) || "Komponen " + (components.length + 1),
				col: i,
				weight: parseWeight(h),
			});
		}
	});
	if (!components.length) {
		alert("Tidak ada kolom skor (Nilai ...) yang dikenali di file ini.");
		return false;
	}
	// Tanpa bobot di judul → bagi rata
	if (components.reduce((x, c) => x + c.weight, 0) === 0) {
		const w = 100 / components.length;
		components.forEach((c) => (c.weight = Math.round(w * 100) / 100));
	}

	students = rawRows
		.map((r) => ({
			nim: String(r[cNim] ?? "").trim(),
			name: String(r[cNama] ?? "").trim(),
			scores: Object.fromEntries(
				components.map((c) => [c.id, parseScore(r[c.col])]),
			),
		}))
		.filter((s) => s.nim || s.name);
	const valid = students.filter((s) => /\d{5,}/.test(s.nim));
	if (valid.length && valid.length < students.length) students = valid;
	if (!students.length) {
		alert("Tidak ada data mahasiswa ditemukan.");
		return false;
	}
	persist();
	renderAll();
	["chartCard", "distCard", "exportCard"].forEach(
		(id) => ($(id).style.display = "block"),
	);
	return true;
}

$("btnRemoveFile").onclick = () => {
	if (!fullAoa.length && !fileName) {
		alert("Belum ada file yang di-upload.");
		return;
	}
	if (!confirm("Hapus file beserta hasil analisis?")) return;
	fullAoa = [];
	rawHeader = [];
	rawRows = [];
	rowMap = [];
	fileName = "";
	components = [];
	students = [];
	$("fileInput").value = "";
	localStorage.removeItem("obe-analisis");
	["chartCard", "distCard", "exportCard"].forEach(
		(id) => ($(id).style.display = "none"),
	);
	$("fileInfo").textContent =
		"Belum ada file. Data tersimpan otomatis di browser (localStorage).";
};

// ---------- Contoh DDKB (otomatis saat pertama buka) ----------
function loadSample() {
	if (!DDKB_SEED) return false;
	fileName = DDKB_SEED.fileName;
	const W = DDKB_SEED.header.length;
	const aoa = [
		[...DDKB_SEED.header],
		...DDKB_SEED.students.map((s) => {
			const r = new Array(W).fill("");
			r[0] = s.nim;
			r[1] = s.name;
			DDKB_SEED.components.forEach((c, k) => {
				r[c.col] = s.scores[k] == null ? "" : s.scores[k];
			});
			return r;
		}),
	];
	if (!autoApply(aoa)) return false;
	$("fileInfo").textContent =
		"File: " + fileName + " (contoh) | " + students.length + " mahasiswa.";
	return true;
}

// ---------- Grafik (skor angka / grade huruf) ----------
$("segHuruf").onclick = () => {
	graphMode = "huruf";
	persist();
	renderGraph();
};
$("segAngka").onclick = () => {
	graphMode = "angka";
	persist();
	renderGraph();
};

function angkaBins() {
	// Bin kontinu tanpa celah: [lo, lo+10), terakhir [90, 100]
	const bins = [];
	for (let b = 0; b < 10; b++) {
		const lo = b * 10;
		bins.push({
			label: lo + "–" + (b === 9 ? 100 : lo + 9),
			lo,
			last: b === 9,
		});
	}
	return bins;
}

function renderGraph() {
	$("segHuruf").classList.toggle("seg-active", graphMode === "huruf");
	$("segAngka").classList.toggle("seg-active", graphMode === "angka");
	const finals = students.map(finalOf);
	const rated = finals.filter((v) => v != null);
	const avg = avgOf(finals);
	$("chartSummary").textContent =
		"File: " +
		(fileName || "-") +
		" • N = " +
		students.length +
		" • Rata-rata akhir: " +
		fmt(avg);
	const box = $("chartBox");
	box.innerHTML = "";
	if (graphMode === "huruf") {
		LADDER.forEach((g) => {
			const n = finals.filter((f) => gradeOf(f) === g.label).length;
			const p = rated.length ? (n / rated.length) * 100 : 0;
			box.appendChild(
				barRow(g.label, p, n + " mhs (" + p.toFixed(1) + "%)", true),
			);
		});
	} else {
		angkaBins().forEach((b) => {
			const n = finals.filter(
				(f) => f != null && f >= b.lo && (b.last ? f <= 100 : f < b.lo + 10),
			).length;
			const p = rated.length ? (n / rated.length) * 100 : 0;
			box.appendChild(
				barRow(b.label, p, n + " mhs (" + p.toFixed(1) + "%)", true),
			);
		});
	}
}

function barRow(label, widthPct, text, good) {
	const d = document.createElement("div");
	d.className = "bar-row";
	const w = Math.max(0, Math.min(100, widthPct));
	d.innerHTML =
		'<span class="bar-label">' +
		escapeHtml(label) +
		"</span>" +
		'<span class="bar-track"><span class="bar-fill ' +
		(good ? "" : "bar-fill-bad") +
		'" style="width:' +
		w.toFixed(1) +
		'%"></span></span>' +
		'<span class="bar-val">' +
		escapeHtml(text) +
		"</span>";
	return d;
}

// ---------- Distribusi grade ----------
function renderDist() {
	const finals = students.map(finalOf);
	const rated = finals.filter((v) => v != null);
	const tb = $("distBody");
	tb.innerHTML = "";
	LADDER.forEach((g, i) => {
		const n = finals.filter((f) => gradeOf(f) === g.label).length;
		const p = rated.length ? (n / rated.length) * 100 : null;
		const tr = document.createElement("tr");
		tr.innerHTML =
			"<td><b>" +
			escapeHtml(g.label) +
			"</b></td>" +
			"<td>" +
			escapeHtml(rentang(i)) +
			"</td>" +
			"<td>" +
			n +
			"</td>" +
			"<td>" +
			(p == null ? "-" : p.toFixed(1) + "%") +
			"</td>";
		tb.appendChild(tr);
	});
	const trT = document.createElement("tr");
	trT.className = "sum-row";
	trT.innerHTML =
		"<td colspan='2'><b>Total</b></td><td><b>" +
		rated.length +
		"</b></td>" +
		"<td><b>" +
		(rated.length ? "100.0%" : "-") +
		"</b></td>";
	tb.appendChild(trT);
}

function renderAll() {
	renderGraph();
	renderDist();
}

// ---------- Export ----------
$("btnExport").onclick = async () => {
	if (typeof ExcelJS === "undefined") {
		alert(
			"Library ExcelJS belum termuat (butuh internet sekali saat buka halaman).",
		);
		return;
	}
	if (!students.length) {
		alert("Belum ada data.");
		return;
	}
	const wb = new ExcelJS.Workbook();
	// Sheet 1: Rekap
	const ws = wb.addWorksheet("Rekap Analisis");
	const head = ["No", "NIM", "Nama", "Nilai Akhir Angka", "Nilai Akhir Huruf"];
	head.forEach((h, i) => (ws.getRow(1).getCell(i + 1).value = h));
	styleHeader(ws, 1, head.length);
	students.forEach((s, i) => {
		const r = i + 2;
		const fin = finalOf(s);
		ws.getRow(r).getCell(1).value = i + 1;
		ws.getRow(r).getCell(2).value = s.nim;
		ws.getRow(r).getCell(3).value = s.name;
		ws.getRow(r).getCell(4).value = fin;
		ws.getRow(r).getCell(4).numFmt = "0.00";
		ws.getRow(r).getCell(5).value = gradeOf(fin);
	});
	ws.columns = [
		{ width: 5 },
		{ width: 16 },
		{ width: 32 },
		{ width: 15 },
		{ width: 12 },
	];
	ws.views = [{ state: "frozen", ySplit: 1 }];
	// Sheet 2: Distribusi (seperti Sheet 2 file OBE)
	const wa = wb.addWorksheet("Distribusi Grade");
	wa.getRow(1).getCell(1).value = "DISTRIBUSI GRADE NILAI AKHIR";
	wa.mergeCells(1, 1, 1, 4);
	wa.getRow(1).getCell(1).font = { bold: true, size: 14 };
	wa.getRow(2).getCell(1).value =
		"File: " +
		fileName +
		" • N = " +
		students.length +
		" • Rata-rata: " +
		fmt(avgOf(students.map(finalOf)));
	const h2 = ["Grade", "Rentang Skor", "Jumlah Mahasiswa", "Persentase (%)"];
	h2.forEach((h, i) => (wa.getRow(4).getCell(i + 1).value = h));
	styleHeader(wa, 4, h2.length);
	const finals = students.map(finalOf);
	const rated = finals.filter((v) => v != null);
	LADDER.forEach((g, i) => {
		const r = 5 + i;
		const n = finals.filter((f) => gradeOf(f) === g.label).length;
		wa.getRow(r).getCell(1).value = g.label;
		wa.getRow(r).getCell(2).value = rentang(i);
		wa.getRow(r).getCell(3).value = n;
		const cell = wa.getRow(r).getCell(4);
		cell.value = rated.length ? n / rated.length : 0;
		cell.numFmt = "0%";
	});
	const rT = 5 + LADDER.length;
	wa.getRow(rT).getCell(1).value = "Total";
	wa.getRow(rT).getCell(1).font = { bold: true };
	wa.getRow(rT).getCell(3).value = rated.length;
	const cT = wa.getRow(rT).getCell(4);
	cT.value = rated.length ? 1 : 0;
	cT.numFmt = "0%";
	wa.columns = [{ width: 10 }, { width: 16 }, { width: 18 }, { width: 15 }];

	const buf = await wb.xlsx.writeBuffer();
	const a = document.createElement("a");
	a.href = URL.createObjectURL(
		new Blob([buf], {
			type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
		}),
	);
	a.download =
		(fileName ? fileName.replace(/\.[^.]+$/, "") : "analisis") +
		"-analisis.xlsx";
	document.body.appendChild(a);
	a.click();
	setTimeout(() => {
		URL.revokeObjectURL(a.href);
		a.remove();
	}, 800);
};

function styleHeader(ws, row, ncols) {
	for (let i = 1; i <= ncols; i++) {
		const c = ws.getRow(row).getCell(i);
		c.font = { bold: true, color: { argb: "FFFFFFFF" } };
		c.fill = {
			type: "pattern",
			pattern: "solid",
			fgColor: { argb: "FF1A1A2E" },
		};
		c.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
	}
	ws.getRow(row).height = 30;
}

// ---------- Simpan / muat ----------
function persist() {
	try {
		localStorage.setItem(
			"obe-analisis",
			JSON.stringify({
				fileName,
				fullAoa,
				headerRowIdx,
				rawHeader,
				cNama,
				cNim,
				components,
				students,
				graphMode,
			}),
		);
	} catch (e) {}
}
function showCards() {
	["chartCard", "distCard", "exportCard"].forEach(
		(id) => ($(id).style.display = "block"),
	);
}
(function restore() {
	try {
		const d = JSON.parse(localStorage.getItem("obe-analisis") || "null");
		if (!d || !d.fullAoa || !d.components || !d.students) return false;
		if (!d.students.length || !d.components.length) return false;
		fileName = d.fileName || "";
		fullAoa = d.fullAoa || [];
		headerRowIdx = d.headerRowIdx || 0;
		rawHeader = d.rawHeader || [];
		cNama = d.cNama;
		cNim = d.cNim;
		components = d.components;
		students = d.students;
		graphMode = d.graphMode || "huruf";
		$("fileInfo").textContent =
			"Data dipulihkan dari sesi sebelumnya (" +
			students.length +
			" mahasiswa).";
		renderAll();
		showCards();
		return true;
	} catch (e) {
		return false;
	}
})();

// ---------- Tema ----------
function initTheme() {
	let t = "light";
	try {
		t = localStorage.getItem("nilai-mhs-theme") || "light";
	} catch (_) {}
	document.documentElement.dataset.theme = t;
	const b = $("btnTheme");
	if (b) b.textContent = t === "dark" ? "☀️" : "🌙";
}
$("btnTheme").onclick = () => {
	const t =
		document.documentElement.dataset.theme === "dark" ? "light" : "dark";
	document.documentElement.dataset.theme = t;
	try {
		localStorage.setItem("nilai-mhs-theme", t);
	} catch (_) {}
	$("btnTheme").textContent = t === "dark" ? "☀️" : "🌙";
};
initTheme();

// Tanpa sesi tersimpan → langsung muat contoh DDKB agar grafik terlihat
if (!students.length) {
	loadSample();
}
