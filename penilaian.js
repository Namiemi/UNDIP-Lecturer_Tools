let fullAoa = []; // seluruh isi sheet asli
let headerRowIdx = 0; // index baris header di fullAoa
let metaRows = []; // baris info di atas header (Mata Kuliah, dll) — dipertahankan saat export
let rawHeader = []; // header asli excel (baris judul kolom yang benar)
let rawRows = []; // baris data asli (array of arrays)
let components = []; // [{id,label,col}] col: index kolom Excel, -1 = kolom baru, "" = belum dipilih (UI)
let compSeq = 0;
let students = []; // [{name,nim,grades:{compId:val},row,r}]
let selectedIdx = -1;
let fileName = "";
let origWb = null; // workbook asli (untuk pertahankan lebar kolom & merge saat export)
let origSheetName = "";
let origCols = null; // salinan ws['!cols']
let origMerges = null; // salinan ws['!merges']
let origBuffer = null; // file asli utuh (untuk export via ExcelJS: preserve font/format/proteksi)
let origB64 = null; // versi base64 origBuffer untuk localStorage
let rowMap = []; // rawRows[i] -> index baris di fullAoa (untuk tulis balik tepat ke sel asli)
let addedCols = {}; // judul kolom baru -> index kolom (dibuat saat export bedah)
let tbl = {
	sortKey: null,
	sortDir: 1,
	filter: "all",
	page: 0,
	perPage: 20,
}; // state tabel canggih
let undoStack = []; // [{label, changes:[{i,t,prev,next}]}] t = component id

const $ = (id) => document.getElementById(id);

function newCompId() {
	compSeq += 1;
	return "k" + Date.now().toString(36) + "_" + compSeq;
}
function compLabel(id) {
	const c = components.find((c) => c.id === id);
	return c ? c.label : id;
}
function getGrade(s, id) {
	if (!s || !s.grades) return null;
	const v = s.grades[id];
	return v == null ? null : v;
}

// ---------- Upload ----------
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
			const wb = XLSX.read(e.target.result, {
				type: "array",
				cellStyles: true,
			});
			const sheetName = wb.SheetNames[0];
			const ws = wb.Sheets[sheetName];
			const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
			if (aoa.length < 2) {
				alert("Excel kosong / tidak ada data.");
				return;
			}
			// simpan workbook asli agar lebar kolom & merge bisa dipertahankan saat download
			origWb = wb;
			origSheetName = sheetName;
			origCols = ws["!cols"] ? JSON.parse(JSON.stringify(ws["!cols"])) : null;
			origMerges = ws["!merges"]
				? JSON.parse(JSON.stringify(ws["!merges"]))
				: null;
			// simpan file mentah untuk export presisi via ExcelJS (font, format angka, proteksi tetap)
			origBuffer = e.target.result;
			try {
				origB64 = bufToB64(origBuffer);
			} catch (_) {
				origB64 = null;
			}
			loadRaw(aoa);
			$("fileInfo").textContent =
				"File: " +
				fileName +
				" | Sheet: " +
				sheetName +
				" | " +
				rawRows.length +
				" baris.";
		} catch (err) {
			alert("Gagal membaca file: " + err.message);
		}
	};
	reader.readAsArrayBuffer(file);
}

function loadRaw(aoa) {
	fullAoa = aoa;
	headerRowIdx = detectHeaderRow(aoa);
	applyHeaderRow(headerRowIdx);
	buildHeaderRowSel();
	buildMappingUI();
	$("mapCard").style.display = "block";
	updateStepper();
}

// Cari baris yang paling mirip header: harus mengandung NIM + Nama dalam satu baris.
// Ini melewati baris info seperti "Mata Kuliah : ...", "Tahun Ajaran : ...", dll.
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
		// baris info seperti "mata kuliah", "tahun ajaran", "semester :" harus dikurangi nilainya
		if (joined.includes("mata kuliah") && !hasNim) score -= 20;
		if (joined.includes("tahun ajaran") && !hasNim) score -= 20;
		if (cells.filter((c) => c !== "").length < 2) score -= 10; // baris kosong
		if (score > bestScore) {
			bestScore = score;
			best = r;
		}
	}
	// jika tidak ada yang meyakinkan (skor rendah), fallback ke baris 0
	if (bestScore < 10) return 0;
	return best;
}

function applyHeaderRow(idx) {
	headerRowIdx = idx;
	rawHeader = (fullAoa[idx] || []).map((h) => String(h ?? "").trim());
	metaRows = fullAoa.slice(0, idx);
	rawRows = [];
	rowMap = [];
	addedCols = {};
	fullAoa.slice(idx + 1).forEach((r, k) => {
		if (r.some((c) => String(c ?? "").trim() !== "")) {
			rowMap.push(idx + 1 + k);
			rawRows.push(r);
		}
	});
	const prev = $("headerPreview");
	if (prev)
		prev.textContent =
			"Baris " +
			(idx + 1) +
			" dipakai sebagai header: " +
			rawHeader.filter(Boolean).slice(0, 6).join(" | ") +
			(rawHeader.filter(Boolean).length > 6 ? " ..." : "");
}

function buildHeaderRowSel() {
	const sel = $("headerRowSel");
	sel.innerHTML = "";
	const limit = Math.min(fullAoa.length, 20);
	for (let r = 0; r < limit; r++) {
		const o = document.createElement("option");
		o.value = r;
		const preview = fullAoa[r]
			.map((c) => String(c ?? "").trim())
			.filter(Boolean)
			.slice(0, 3)
			.join(" | ")
			.slice(0, 60);
		o.textContent =
			"Baris " + (r + 1) + (preview ? " — " + preview : " — (kosong)");
		sel.appendChild(o);
	}
	sel.value = headerRowIdx;
	sel.onchange = () => {
		applyHeaderRow(parseInt(sel.value));
		buildMappingUI();
	};
}

// ---------- Mapping (Nama/NIM otomatis; komponen nilai dipilih manual) ----------
function guessIndex(patterns) {
	for (let i = 0; i < rawHeader.length; i++) {
		const h = rawHeader[i].toLowerCase();
		if (patterns.some((p) => h.includes(p))) return i;
	}
	return -1;
}

function fillSelect(sel, includeNew) {
	sel.innerHTML = "";
	const ph = document.createElement("option");
	ph.value = "";
	ph.textContent = "— Pilih kolom —";
	sel.appendChild(ph);
	rawHeader.forEach((h, i) => {
		const o = document.createElement("option");
		o.value = i;
		const label = h || "(Kolom kosong " + (i + 1) + ")";
		// tampilkan nama lengkap via title, potong kalau terlalu panjang
		o.textContent =
			"Kol " + (i + 1) + ": " + (label.length > 45 ? label.slice(0, 45) + "…" : label);
		o.title = label;
		sel.appendChild(o);
	});
	if (includeNew) {
		const o = document.createElement("option");
		o.value = "__new__";
		o.textContent = "➕ Buat kolom baru";
		sel.appendChild(o);
	}
	sel.value = "";
}

function buildMappingUI() {
	fillSelect($("mapNama"), false);
	fillSelect($("mapNim"), false);
	// Kolom Nama & NIM tetap ditebak otomatis; komponen nilai tidak.
	const iNama = guessIndex(["nama mahasiswa", "nama", "name"]);
	const iNim = guessIndex(["nim"]);
	if (iNama >= 0) $("mapNama").value = iNama;
	if (iNim >= 0) $("mapNim").value = iNim;
	// Tepat 3 komponen: pertahankan yang sudah ada, tambah baris kosong bila kurang.
	while (components.length < 3) {
		components.push({ id: newCompId(), label: "", col: "" });
	}
	if (components.length > 3) components = components.slice(0, 3);
	// Header berubah → indeks kolom lama bisa tidak valid lagi.
	components.forEach((c) => {
		if (c.col !== "__new__" && c.col !== "" && c.col != null) {
			const n = parseInt(c.col);
			if (isNaN(n) || n < 0 || n >= rawHeader.length) c.col = "";
		}
	});
	renderCompRows();
}

// Nama komponen otomatis mengikuti judul kolom Excel yang dipilih
// (cth: "Nilai UTS (0,00-100,00) (15%)" → "UTS").
// Tidak menimpa bila pengguna sudah mengetik nama sendiri.
function suggestName(header) {
	let s = String(header ?? "").split("(")[0].trim();
	s = s.replace(/^nilai\s+/i, "").trim();
	return s;
}

function renderCompRows() {
	const list = $("gradeMapList");
	if (!list) return;
	list.innerHTML = "";
	const hints = ["UTS", "UAS", "Tugas"];
	components.forEach((c, idx) => {
		const row = document.createElement("div");
		row.className = "comp-row";

		const tag = document.createElement("span");
		tag.className = "comp-tag";
		tag.textContent = "Kolom " + (idx + 1);

		const inp = document.createElement("input");
		inp.type = "text";
		inp.className = "comp-name";
		inp.placeholder =
			"Nama komponen " + (idx + 1) + " (cth: " + (hints[idx] || "Kuis") + ")";
		inp.value = c.label || "";
		inp.oninput = () => {
			components[idx].label = inp.value;
			components[idx].autoLabel = false; // ketikan sendiri → jangan ditimpa otomatis
		};

		const sel = document.createElement("select");
		sel.className = "comp-col";
		fillSelect(sel, true);
		sel.value = c.col === -1 ? "__new__" : (c.col ?? "");
		sel.onchange = () => {
			components[idx].col = sel.value;
			const cc = components[idx];
			const v = sel.value;
			if (
				v !== "" &&
				v !== "__new__" &&
				(cc.autoLabel || !String(cc.label || "").trim())
			) {
				const name = suggestName(rawHeader[parseInt(v)] || "");
				if (name) {
					cc.label = name;
					inp.value = name;
					cc.autoLabel = true;
				}
			}
		};

		row.appendChild(tag);
		row.appendChild(inp);
		row.appendChild(sel);
		list.appendChild(row);
	});
}

$("btnApplyMap").onclick = () => {
	const namaVal = $("mapNama").value;
	const nimVal = $("mapNim").value;
	if (namaVal === "" || nimVal === "") {
		alert("Pilih dulu kolom Nama dan kolom NIM.");
		return;
	}
	const cNama = parseInt(namaVal);
	const cNim = parseInt(nimVal);
	if (cNama === cNim) {
		alert("Kolom Nama dan NIM tidak boleh sama!");
		return;
	}
	const rows = [...document.querySelectorAll("#gradeMapList .comp-row")];
	if (!rows.length) {
		alert("Tambah dulu minimal 1 komponen nilai.");
		return;
	}
	const comps = rows.map((row) => ({
		id: components[[...row.parentNode.children].indexOf(row)]?.id || newCompId(),
		label: row.querySelector(".comp-name").value.trim(),
		raw: row.querySelector(".comp-col").value,
	}));
	// Validasi label
	if (comps.some((c) => !c.label)) {
		alert("Semua komponen nilai harus diberi nama (cth: Tugas, Kuis, UTS, UAS).");
		return;
	}
	const lower = comps.map((c) => c.label.toLowerCase());
	if (new Set(lower).size !== lower.length) {
		alert("Nama komponen tidak boleh sama (periksa duplikat).");
		return;
	}
	// Validasi kolom
	if (comps.some((c) => c.raw === "")) {
		alert("Setiap komponen harus dipetakan: pilih kolom Excel atau 'Buat kolom baru'.");
		return;
	}
	const norm = comps.map((c) => ({
		id: c.id,
		label: c.label,
		col: c.raw === "__new__" ? -1 : parseInt(c.raw),
	}));
	const usedExisting = norm.filter((c) => c.col >= 0).map((c) => c.col);
	if (usedExisting.some((c) => c === cNama || c === cNim)) {
		alert("Kolom komponen nilai tidak boleh sama dengan kolom Nama / NIM.");
		return;
	}
	if (new Set(usedExisting).size !== usedExisting.length) {
		alert("Dua komponen tidak boleh memakai kolom Excel yang sama.");
		return;
	}
	components = norm;
	students = rawRows
		.map((r, idx) => {
			const grades = {};
			components.forEach((c) => {
				grades[c.id] = c.col >= 0 ? parseScore(r[c.col]) : null;
			});
			return {
				name: String(r[cNama] ?? "").trim(),
				nim: String(r[cNim] ?? "").trim(),
				grades,
				row: idx,
				r: rowMap[idx],
			};
		})
		.filter((s) => s.name || s.nim);
	if (!students.length) {
		alert("Tidak ada data mahasiswa ditemukan.");
		return;
	}
	// Abaikan baris kaki/footer (tanda tangan, kode verifikasi, dsb): NIM asli selalu mengandung deret angka.
	// Hanya aktif bila ada ≥1 baris ber-NIM valid, supaya tidak menghapus semua saat mapping salah pilih.
	const valid = students.filter((s) => /\d{5,}/.test(s.nim));
	let dropped = 0;
	if (valid.length && valid.length < students.length) {
		dropped = students.length - valid.length;
		students = valid;
	}
	const mi = $("mapInfo");
	if (mi)
		mi.textContent =
			students.length +
			" mahasiswa terdeteksi • " +
			components.length +
			" komponen (" +
			components.map((c) => c.label).join(", ") +
			")" +
			(dropped
				? " • " + dropped + " baris non-mahasiswa (footer/kode) diabaikan"
				: "") +
			".";
	tbl.page = 0;
	tbl.filter = "all";
	tbl.sortKey = null;
	undoStack = [];
	updateUndoBtn();
	rebuildDynamicUI();
	persist();
	renderAll();
	$("gradeCard").style.display = "block";
	$("tableCard").style.display = "block";
	$("gradeCard").scrollIntoView({ behavior: "smooth" });
};

function parseScore(v) {
	if (v === "" || v == null) return null;
	const n = parseFloat(String(v).replace(",", "."));
	return isNaN(n) ? null : n;
}

// ---------- UI dinamis mengikuti komponen ----------
function rebuildDynamicUI() {
	// Dropdown tipe nilai
	const gt = $("gradeType");
	gt.innerHTML = "";
	components.forEach((c) => {
		const o = document.createElement("option");
		o.value = c.id;
		o.textContent = c.label;
		gt.appendChild(o);
	});
	// Header tabel
	const hr = $("theadRow");
	hr.innerHTML = "";
	const thNo = document.createElement("th");
	thNo.textContent = "No";
	hr.appendChild(thNo);
	hr.appendChild(makeSortTh("name", "Nama"));
	hr.appendChild(makeSortTh("nim", "NIM"));
	components.forEach((c) => {
		hr.appendChild(makeSortTh(c.id, c.label));
	});
	const thAct = document.createElement("th");
	thAct.textContent = "Aksi";
	hr.appendChild(thAct);
	wireSortHeaders();
	// Filter
	const f = $("tblFilter");
	f.innerHTML = "";
	const addOpt = (val, text) => {
		const o = document.createElement("option");
		o.value = val;
		o.textContent = text;
		f.appendChild(o);
	};
	addOpt("all", "Filter: Semua");
	components.forEach((c) => {
		addOpt("need-" + c.id, "Belum ada " + c.label);
	});
	addOpt("todo", "Belum lengkap");
	addOpt(
		"done",
		"Lengkap (" + components.map((c) => c.label).join("+") + ")",
	);
	if (!["all", "todo", "done", ...components.map((c) => "need-" + c.id)].includes(tbl.filter)) {
		tbl.filter = "all";
	}
	f.value = tbl.filter;
	// Tombol hapus per komponen
	const cb = $("clearBtns");
	cb.innerHTML = "";
	components.forEach((c) => {
		const wrap = document.createElement("div");
		wrap.className = "col";
		const b = document.createElement("button");
		b.className = "btn btn-red";
		b.style.width = "100%";
		b.textContent = "🗑 Hapus " + c.label;
		b.title = "Kosongkan semua nilai di komponen " + c.label;
		b.onclick = () => clearColumn(c.id);
		wrap.appendChild(b);
		cb.appendChild(wrap);
	});
}

function makeSortTh(key, label) {
	const th = document.createElement("th");
	th.className = "sortable";
	th.dataset.sort = key;
	th.innerHTML = escapeHtml(label) + ' <span class="sort-arrow"></span>';
	return th;
}

// ---------- Search + grade (keyboard friendly, tanpa mouse) ----------
let searchList = []; // hasil filter terakhir [{...s,i}]
let kbIdx = -1; // index aktif di searchList untuk navigasi keyboard

$("searchInput").addEventListener("input", updateSearch);
$("searchInput").addEventListener("keydown", (e) => {
	if (e.key === "ArrowDown") {
		e.preventDefault();
		if (searchList.length) {
			kbIdx = Math.min(kbIdx + 1, searchList.length - 1);
			paintKb();
		}
	} else if (e.key === "ArrowUp") {
		e.preventDefault();
		if (searchList.length) {
			kbIdx = Math.max(kbIdx - 1, 0);
			paintKb();
		}
	} else if (e.key === "Enter") {
		e.preventDefault();
		if (searchList.length) {
			const pick = kbIdx >= 0 ? searchList[kbIdx] : searchList[0];
			selectStudent(pick.i);
			// langsung lompat ke input skor → bisa ketik nilai + Enter, tanpa mouse sama sekali
			$("gradeScore").focus();
		}
	} else if (e.key === "Escape") {
		e.target.value = "";
		updateSearch();
	}
});

// Enter di kolom skor = simpan. Setelah simpan, fokus balik ke pencarian untuk input berikutnya.
$("gradeScore").addEventListener("keydown", (e) => {
	if (e.key === "Enter") {
		e.preventDefault();
		$("btnSaveGrade").click();
	}
});

// Shortcut global: "/" atau Ctrl+K fokus ke pencarian (kecuali sedang mengetik di input lain)
document.addEventListener("keydown", (e) => {
	const tag = (document.activeElement?.tagName || "").toLowerCase();
	const typing =
		tag === "input" || tag === "select" || tag === "textarea";
	if (
		(e.key === "/" && !typing) ||
		((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === "k")
	) {
		e.preventDefault();
		$("gradeCard").style.display !== "none" && $("searchInput").focus();
	}
});

function updateSearch() {
	const q = $("searchInput").value.trim().toLowerCase();
	const box = $("searchResults");
	box.innerHTML = "";
	searchList = [];
	kbIdx = -1;
	if (!q) return;
	searchList = students
		.map((s, i) => ({ ...s, i }))
		.filter(
			(s) =>
				s.name.toLowerCase().includes(q) ||
				s.nim.toLowerCase().includes(q),
		)
		.slice(0, 20);
	if (!searchList.length) {
		box.innerHTML = "<p class='hint'>Tidak ditemukan.</p>";
		return;
	}
	kbIdx = 0; // default: hasil pertama langsung aktif → tinggal Enter
	searchList.forEach((s, pos) => {
		const d = document.createElement("div");
		d.className =
			"result-item" +
			(s.i === selectedIdx ? " active" : "") +
			(pos === kbIdx ? " kb-active" : "");
		d.innerHTML =
			"<span><b>" +
			hi(s.name, q) +
			"</b><br><small>" +
			hi(s.nim, q) +
			"</small></span><span class='kbd-hint'>" +
			(pos === 0 ? "Enter ⏎" : "↑↓+Enter") +
			"</span>";
		d.onclick = () => {
			selectStudent(s.i);
		};
		d.onmousemove = () => {
			if (kbIdx !== pos) {
				kbIdx = pos;
				paintKb();
			}
		};
		box.appendChild(d);
	});
}

// Tandai cocoknya pencarian dengan <mark> (q sudah lowercase)
function hi(text, q) {
	const e = escapeHtml(text || "-");
	if (!q) return e;
	const i = e.toLowerCase().indexOf(q);
	if (i < 0) return e;
	return (
		e.slice(0, i) +
		"<mark>" +
		e.slice(i, i + q.length) +
		"</mark>" +
		e.slice(i + q.length)
	);
}

function paintKb() {
	const box = $("searchResults");
	[...box.children].forEach((el, pos) => {
		el.classList.toggle("kb-active", pos === kbIdx);
		const hint = el.querySelector(".kbd-hint");
		if (hint) hint.textContent = pos === kbIdx ? "Enter ⏎" : "↑↓+Enter";
	});
	// scroll item aktif ke dalam view
	const active = box.children[kbIdx];
	if (active) active.scrollIntoView({ block: "nearest" });
}

function selectStudent(i) {
	selectedIdx = i;
	const s = students[i];
	$("selectedInfo").textContent =
		"Dipilih: " + s.name + " (" + s.nim + ")";
	$("foundBox").style.display = "block";
	$("foundName").textContent = s.name || "-";
	$("foundNim").textContent = "NIM: " + (s.nim || "-");
	const fg = $("foundGrades");
	fg.innerHTML = "";
	components.forEach((c) => {
		const v = getGrade(s, c.id);
		const sp = document.createElement("span");
		sp.className = "badge badge-gen";
		sp.textContent = c.label + ": " + (v == null ? "-" : v);
		fg.appendChild(sp);
	});
	// lompat ke halaman tabel yang memuat baris ini
	const pos = tableView().indexOf(i);
	if (pos >= 0) tbl.page = Math.floor(pos / tbl.perPage);
	renderTable(); // refresh highlight
	paintKb();
}

function saveGrade() {
	if (selectedIdx < 0) {
		alert("Pilih dulu mahasiswa (ketik Nama/NIM, lalu tekan Enter).");
		$("searchInput").focus();
		return;
	}
	if (!components.length) {
		alert("Belum ada komponen nilai. Terapkan pemetaan dulu.");
		return;
	}
	const type = $("gradeType").value;
	if (!components.some((c) => c.id === type)) {
		alert("Komponen nilai tidak valid.");
		return;
	}
	const val = parseFloat(String($("gradeScore").value).replace(",", "."));
	if (isNaN(val) || val < 0 || val > 100) {
		alert("Skor harus angka 0–100.");
		$("gradeScore").focus();
		return;
	}
	pushUndo(
		"nilai " +
			compLabel(type) +
			" " +
			(students[selectedIdx].name || students[selectedIdx].nim),
		[
			{
				i: selectedIdx,
				t: type,
				prev: getGrade(students[selectedIdx], type),
				next: val,
			},
		],
	);
	students[selectedIdx].grades[type] = val;
	$("gradeScore").value = "";
	persist();
	renderAll();
	selectStudent(selectedIdx);
	updateSearch(); // refresh daftar (skor berubah)
	// QoL: fokus balik ke pencarian + blok teks → langsung ketik NIM berikutnya
	$("searchInput").focus();
	$("searchInput").select();
}
$("btnSaveGrade").onclick = saveGrade;

// Alt+1..9 = ganti komponen nilai tanpa mouse
document.addEventListener("keydown", (e) => {
	if (e.altKey && /^[1-9]$/.test(e.key)) {
		const idx = parseInt(e.key, 10) - 1;
		if (components[idx] && $("gradeType")) {
			e.preventDefault();
			$("gradeType").value = components[idx].id;
		}
	}
});

function escapeHtml(s) {
	return String(s).replace(
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

// ---------- Render + stepper progres ----------
function renderAll() {
	renderTable();
	renderStats();
	updateStepper();
}

function setStep(n, st) {
	const el = document.querySelector('.step[data-s="' + n + '"]');
	if (el) el.className = "step" + (st ? " " + st : "");
}
function updateStepper() {
	const hasFile = fullAoa.length > 0;
	const hasStudents = students.length > 0;
	const allDone =
		hasStudents &&
		components.length > 0 &&
		students.every((s) => components.every((c) => getGrade(s, c.id) != null));
	setStep(1, hasFile ? "done" : "active");
	setStep(2, hasStudents ? "done" : hasFile ? "active" : "");
	setStep(3, allDone ? "done" : hasStudents ? "active" : "");
	setStep(4, hasStudents ? (allDone ? "done" : "active") : "");
}

// ---------- Undo (tumpuk perubahan nilai) ----------
function pushUndo(label, changes) {
	if (!changes || !changes.length) return;
	undoStack.push({ label, changes });
	if (undoStack.length > 100) undoStack.shift();
	updateUndoBtn();
}
function updateUndoBtn() {
	const b = $("btnUndo");
	if (!b) return;
	b.disabled = !undoStack.length;
	b.textContent = undoStack.length
		? "↩ Urungkan (" + undoStack.length + ")"
		: "↩ Urungkan";
	b.title = undoStack.length
		? "Terakhir: " + undoStack[undoStack.length - 1].label + " • Ctrl+Z"
		: "Belum ada perubahan";
}
function doUndo() {
	const e = undoStack.pop();
	if (!e) return;
	e.changes.forEach((c) => {
		if (students[c.i]) {
			if (!students[c.i].grades) students[c.i].grades = {};
			students[c.i].grades[c.t] = c.prev;
		}
	});
	persist();
	renderAll();
	if (selectedIdx >= 0) selectStudent(selectedIdx);
	updateSearch();
	updateUndoBtn();
	const ss = $("saveState");
	if (ss) ss.textContent = "↩ Dibatalkan: " + e.label;
}
$("btnUndo").onclick = doUndo;
// Ctrl+Z global — kecuali sedang mengetik di input/textarea (biar undo teks browser tetap jalan)
document.addEventListener("keydown", (e) => {
	if ((e.ctrlKey || e.metaKey) && !e.shiftKey && e.key.toLowerCase() === "z") {
		const tag = (document.activeElement?.tagName || "").toLowerCase();
		if (tag === "input" || tag === "textarea") return;
		e.preventDefault();
		doUndo();
	}
});

// ---------- Tema gelap / terang ----------
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

function renderStats() {
	$("statTotal").textContent = students.length;
	const row = $("statRow");
	if (!row) return;
	row.querySelectorAll("[data-stat]").forEach((el) => el.remove());
	components.forEach((c) => {
		const d = document.createElement("div");
		d.setAttribute("data-stat", c.id);
		const b = document.createElement("b");
		b.textContent = students.filter((s) => getGrade(s, c.id) != null).length;
		const sp = document.createElement("span");
		sp.textContent = "Sudah ada " + c.label;
		d.appendChild(b);
		d.appendChild(sp);
		row.appendChild(d);
	});
}

// ---------- Tabel canggih: filter + sortir + pagination + badge warna ----------
function tableView() {
	let idx = students.map((_, i) => i);
	const f = tbl.filter;
	if (f.startsWith("need-")) {
		const id = f.slice(5);
		idx = idx.filter((i) => getGrade(students[i], id) == null);
	} else if (f === "todo")
		idx = idx.filter((i) =>
			components.some((c) => getGrade(students[i], c.id) == null),
		);
	else if (f === "done")
		idx = idx.filter((i) =>
			components.every((c) => getGrade(students[i], c.id) != null),
		);
	if (tbl.sortKey) {
		const k = tbl.sortKey,
			d = tbl.sortDir;
		const val = (i) =>
			k === "name"
				? (students[i].name || "").toLowerCase()
				: k === "nim"
					? String(students[i].nim || "")
					: getGrade(students[i], k);
		idx.sort((a, b) => {
			const va = val(a),
				vb = val(b);
			if (va == null && vb == null) return 0;
			if (va == null) return 1;
			if (vb == null) return -1;
			if (va < vb) return -d;
			if (va > vb) return d;
			return 0;
		});
	}
	return idx;
}

// Warna nilai: ≥80 hijau, 60–79 kuning, <60 merah (batas umum kelulusan)
function gradeClass(v) {
	if (v == null || v === "") return "";
	if (v >= 80) return "g-good";
	if (v >= 60) return "g-mid";
	return "g-low";
}

function renderTable() {
	const tb = $("tbody");
	tb.innerHTML = "";
	if (!components.length) {
		$("pgInfo").textContent = "Terapkan pemetaan dulu untuk menampilkan tabel.";
		return;
	}
	const view = tableView();
	const pages = Math.max(1, Math.ceil(view.length / tbl.perPage));
	if (tbl.page >= pages) tbl.page = pages - 1;
	if (tbl.page < 0) tbl.page = 0;
	const start = tbl.page * tbl.perPage;
	const rows = view.slice(start, start + tbl.perPage);
	rows.forEach((i, pos) => {
		const s = students[i];
		const tr = document.createElement("tr");
		if (i === selectedIdx) tr.style.background = "var(--hover)";
		let html =
			"<td>" +
			(start + pos + 1) +
			"</td>" +
			"<td style='text-align:left'>" +
			escapeHtml(s.name) +
			"</td>" +
			"<td>" +
			escapeHtml(s.nim) +
			"</td>";
		components.forEach((c) => {
			const v = getGrade(s, c.id);
			html +=
				"<td><input class='grade-input " +
				gradeClass(v) +
				"' type='number' min='0' max='100' value='" +
				(v == null ? "" : v) +
				"' data-i='" +
				i +
				"' data-t='" +
				c.id +
				"'></td>";
		});
		html +=
			"<td><button class='btn btn-dark' style='padding:5px 12px;font-size:12px' data-pick='" +
			i +
			"'>Pilih</button></td>";
		tr.innerHTML = html;
		tb.appendChild(tr);
	});
	$("pgInfo").textContent = view.length
		? "Baris " +
			(start + 1) +
			"–" +
			(start + rows.length) +
			" dari " +
			view.length +
			" • Hal " +
			(tbl.page + 1) +
			"/" +
			pages
		: "Tidak ada baris (coba ubah filter).";
	$("pgPrev").disabled = tbl.page <= 0;
	$("pgNext").disabled = tbl.page >= pages - 1;
	document.querySelectorAll("th.sortable").forEach((th) => {
		const a = th.querySelector(".sort-arrow");
		if (a)
			a.textContent =
				th.dataset.sort === tbl.sortKey ? (tbl.sortDir === 1 ? "▲" : "▼") : "";
	});
	tb.querySelectorAll(".grade-input").forEach((inp) => {
		inp.onchange = (e) => {
			const i = +e.target.dataset.i,
				t = e.target.dataset.t;
			const prev = getGrade(students[i], t);
			const v =
				e.target.value.trim() === "" ? null : parseFloat(e.target.value);
			if (v != null && (isNaN(v) || v < 0 || v > 100)) {
				alert("Skor 0–100 saja.");
				renderTable();
				return;
			}
			if (prev === v) {
				renderTable();
				return;
			}
			pushUndo(
				"edit " + compLabel(t) + " " + (students[i].name || students[i].nim),
				[{ i, t, prev, next: v }],
			);
			if (!students[i].grades) students[i].grades = {};
			students[i].grades[t] = v;
			persist();
			renderStats();
			updateStepper();
			if (i === selectedIdx) selectStudent(i);
			else renderTable();
		};
	});
	tb.querySelectorAll("[data-pick]").forEach(
		(b) => (b.onclick = () => selectStudent(+b.dataset.pick)),
	);
}

// Wiring sortir header (dipanggil ulang tiap thead dibangun dinamis)
function wireSortHeaders() {
	document.querySelectorAll("th.sortable").forEach((th) => {
		th.onclick = () => {
			const k = th.dataset.sort;
			if (tbl.sortKey === k) tbl.sortDir *= -1;
			else {
				tbl.sortKey = k;
				tbl.sortDir = 1;
			}
			tbl.page = 0;
			renderTable();
		};
	});
}
$("tblFilter").onchange = (e) => {
	tbl.filter = e.target.value;
	tbl.page = 0;
	renderTable();
};
$("pgPrev").onclick = () => {
	if (tbl.page > 0) {
		tbl.page--;
		renderTable();
	}
};
$("pgNext").onclick = () => {
	tbl.page++;
	renderTable();
};
$("pgSize").onchange = (e) => {
	tbl.perPage = parseInt(e.target.value) || 20;
	tbl.page = 0;
	renderTable();
};

// ---------- Export ----------
// Jalur utama: ExcelJS memodifikasi FILE ASLI → font/size, format angka (90,00),
// lebar kolom, merge, dan proteksi sheet semuanya dipertahankan persis.
$("btnExport").onclick = async () => {
	if (!students.length) return;
	const name =
		(fileName ? fileName.replace(/\.[^.]+$/, "") : "nilai-mahasiswa") +
		"-dinilai.xlsx";
	if (origBuffer && typeof ExcelJS !== "undefined") {
		try {
			await exportViaExcelJS(name);
			return;
		} catch (err) {
			console.warn("Export presisi gagal, pakai mode kompatibilitas:", err);
		}
	}
	exportFallback(name); // tanpa buffer asli (sesi lama) atau file .xls: bangun ulang via SheetJS
};

async function exportViaExcelJS(name) {
	const wb = new ExcelJS.Workbook();
	await wb.xlsx.load(origBuffer.slice(0));
	const ws = wb.getWorksheet(origSheetName) || wb.worksheets[0];
	if (!ws) throw new Error("Sheet tidak ditemukan di file asli.");
	// Petakan tiap komponen ke kolom kerja (pakai kolom lama atau buat baru)
	const works = components.map((c) => ({
		def: c,
		origCol: c.col,
		w: c.col >= 0 ? c.col : ensureXCol(ws, -1, c.label),
	}));
	// Komponen baru sekarang punya kolom tetap → simpan agar konsisten
	works.forEach(({ def, w }) => {
		def.col = w;
	});
	const tpls = works.map(({ w }) => colTemplate(ws, w));
	students.forEach((s) => {
		works.forEach(({ def, origCol, w }, k) => {
			const origVal =
				origCol >= 0 && rawRows[s.row]
					? parseScore(rawRows[s.row][origCol])
					: null;
			if (getGrade(s, def.id) !== origVal)
				setXCell(ws, s.r, w, getGrade(s, def.id), tpls[k]);
		});
	});
	// sinkronkan cache agar export berikutnya & perbandingan tetap akurat
	students.forEach((s) => {
		works.forEach(({ def, w }) => {
			if (!rawRows[s.row]) rawRows[s.row] = [];
			if (!fullAoa[s.r]) fullAoa[s.r] = [];
			const v = getGrade(s, def.id);
			rawRows[s.row][w] = v == null ? "" : v;
			fullAoa[s.r][w] = v == null ? "" : v;
		});
	});
	const buf = await wb.xlsx.writeBuffer();
	downloadBlob(
		new Blob([buf], {
			type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
		}),
		name,
	);
}

// Cari template style kolom: sel berformat pertama di kolom itu (mis. sel kosong yg sudah diformat 0,00).
// Kalau tidak ada, pakai font badan tabel + format 0.00.
function colTemplate(ws, c0) {
	const col = c0 + 1;
	for (const s of students) {
		if (s.r == null) continue;
		const cell = ws.getRow(s.r + 1).getCell(col);
		if (cell.numFmt && cell.numFmt !== "General")
			return {
				numFmt: cell.numFmt,
				font: { ...cell.font },
				alignment: { ...cell.alignment },
				border: { ...cell.border },
				protection: { ...cell.protection },
			};
	}
	return { numFmt: "0.00", font: bodyFont(ws) };
}

function bodyFont(ws) {
	if (students.length && students[0].r != null) {
		const row = ws.getRow(students[0].r + 1);
		for (
			let c = 1;
			c <= rawHeader.length + Object.keys(addedCols).length;
			c++
		) {
			const f = row.getCell(c).font;
			if (f && (f.name || f.size)) return { ...f };
		}
	}
	return { name: "Calibri", size: 11 };
}

function setXCell(ws, r0, c0, val, tpl) {
	const cell = ws.getRow(r0 + 1).getCell(c0 + 1);
	cell.value = val == null || val === "" ? null : val;
	if (tpl.numFmt) cell.numFmt = tpl.numFmt;
	if (tpl.font) cell.font = tpl.font;
	if (tpl.alignment && Object.keys(tpl.alignment).length)
		cell.alignment = tpl.alignment;
	if (tpl.border && Object.keys(tpl.border).length) cell.border = tpl.border;
	if (tpl.protection && Object.keys(tpl.protection).length)
		cell.protection = tpl.protection;
}

function ensureXCol(ws, cur, title) {
	if (cur >= 0) return cur;
	if (addedCols[title] != null) return addedCols[title];
	const nc0 = rawHeader.length + Object.keys(addedCols).length;
	const hr = headerRowIdx + 1,
		nc = nc0 + 1;
	const hcell = ws.getRow(hr).getCell(nc);
	hcell.value = title;
	// tiru style judul tetangga terakhir yang terisi
	for (let c = rawHeader.length; c >= 1; c--) {
		const d = ws.getRow(hr).getCell(c);
		if (d.value != null && d.value !== "") {
			if (d.font) hcell.font = { ...d.font };
			if (d.alignment) hcell.alignment = { ...d.alignment };
			if (d.border) hcell.border = { ...d.border };
			if (d.fill) hcell.fill = { ...d.fill };
			break;
		}
	}
	ws.getColumn(nc).width = 18;
	addedCols[title] = nc0;
	return nc0;
}

function downloadBlob(blob, name) {
	const a = document.createElement("a");
	a.href = URL.createObjectURL(blob);
	a.download = name;
	document.body.appendChild(a);
	a.click();
	setTimeout(() => {
		URL.revokeObjectURL(a.href);
		a.remove();
	}, 800);
}

function bufToB64(buf) {
	const u = new Uint8Array(buf);
	let s = "";
	for (let i = 0; i < u.length; i += 0x8000) {
		s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000));
	}
	return btoa(s);
}
function b64ToBuf(b64) {
	const s = atob(b64);
	const u = new Uint8Array(s.length);
	for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i);
	return u.buffer;
}

// Mode kompatibilitas: bangun ulang via SheetJS (dipakai bila tanpa buffer asli / file .xls).
// Catatan: mode ini tidak mempertahankan font & format angka sedetail ExcelJS.
function exportFallback(name) {
	let header = [...rawHeader];
	const workIdx = components.map((c) => {
		if (c.col >= 0) return c.col;
		header.push(c.label);
		return header.length - 1;
	});
	// Samakan kolom baru agar konsisten dengan export presisi
	components.forEach((c, k) => {
		if (c.col < 0) c.col = workIdx[k];
	});
	const byRow = {};
	students.forEach((s) => (byRow[s.row] = s));
	const dataPart = [header];
	rawRows.forEach((r, idx) => {
		const nr = [...r];
		while (nr.length < header.length) nr.push("");
		const s = byRow[idx];
		if (s) {
			components.forEach((c, k) => {
				const v = getGrade(s, c.id);
				nr[workIdx[k]] = v == null ? "" : v;
			});
		}
		dataPart.push(nr);
	});
	const width = Math.max(header.length, ...metaRows.map((r) => r.length));
	const norm = (r) => {
		const nr = [...r];
		while (nr.length < width) nr.push("");
		return nr;
	};
	const out = [...metaRows.map(norm), ...dataPart.map(norm)];
	const ws2 = XLSX.utils.aoa_to_sheet(out);
	ws2["!cols"] = computeWidths(out, origCols);
	if (origMerges) ws2["!merges"] = origMerges;
	const wb2 = XLSX.utils.book_new();
	XLSX.utils.book_append_sheet(wb2, ws2, origSheetName || "Nilai");
	XLSX.writeFile(wb2, name);
}

function computeWidths(aoa, base) {
	const n = Math.max(...aoa.map((r) => r.length));
	const cols = [];
	for (let c = 0; c < n; c++) {
		if (base && base[c]) {
			cols.push(base[c]);
			continue;
		}
		let m = 10;
		for (let r = 0; r < aoa.length; r++) {
			const len = String(aoa[r][c] ?? "").length;
			if (len > m) m = len;
		}
		cols.push({ wch: Math.min(50, m + 2) });
	}
	return cols;
}

// ---------- Hapus file / hapus nilai per komponen ----------
$("btnRemoveFile").onclick = () => {
	if (fullAoa.length === 0 && !fileName) {
		alert("Belum ada file yang di-upload.");
		return;
	}
	if (!confirm("Hapus file yang sudah di-upload beserta semua data?")) return;
	fullReset();
};

function clearColumn(compId) {
	if (!students.length) {
		alert("Belum ada data.");
		return;
	}
	const label = compLabel(compId);
	const filled = students.filter((s) => getGrade(s, compId) != null).length;
	if (!filled) {
		alert("Tidak ada nilai " + label + " yang perlu dihapus.");
		return;
	}
	if (
		!confirm(
			"Hapus SEMUA nilai " +
				label +
				" (" +
				filled +
				" mahasiswa)? Kolomnya tetap ada, isinya dikosongkan.",
		)
	)
		return;
	pushUndo(
		"hapus nilai " + label + " (" + filled + " mhs)",
		students
			.map((s, i) => ({ i, t: compId, prev: getGrade(s, compId), next: null }))
			.filter((c) => c.prev != null),
	);
	students.forEach((s) => {
		if (!s.grades) s.grades = {};
		s.grades[compId] = null;
	});
	persist();
	renderAll();
	if (selectedIdx >= 0) selectStudent(selectedIdx);
	updateSearch();
}

function fullReset() {
	students = [];
	components = [];
	compSeq = 0;
	rawHeader = [];
	rawRows = [];
	fullAoa = [];
	metaRows = [];
	headerRowIdx = 0;
	selectedIdx = -1;
	searchList = [];
	kbIdx = -1;
	fileName = "";
	origWb = null;
	origSheetName = "";
	origCols = null;
	origMerges = null;
	rowMap = [];
	addedCols = {};
	origBuffer = null;
	origB64 = null;
	tbl = {
		sortKey: null,
		sortDir: 1,
		filter: "all",
		page: 0,
		perPage: 20,
	};
	undoStack = [];
	const tf = $("tblFilter");
	if (tf) {
		tf.innerHTML = "";
		const o = document.createElement("option");
		o.value = "all";
		o.textContent = "Filter: Semua";
		tf.appendChild(o);
		tf.value = "all";
	}
	const hr = $("theadRow");
	if (hr) hr.innerHTML = "";
	const cb = $("clearBtns");
	if (cb) cb.innerHTML = "";
	const gt = $("gradeType");
	if (gt) gt.innerHTML = "";
	$("fileInput").value = "";
	localStorage.removeItem("nilai-mhs");
	["mapCard", "gradeCard", "tableCard"].forEach(
		(id) => ($(id).style.display = "none"),
	);
	$("fileInfo").textContent =
		"Belum ada file. Data tersimpan otomatis di browser (localStorage).";
	$("tbody").innerHTML = "";
	$("searchResults").innerHTML = "";
	$("searchInput").value = "";
	$("gradeScore").value = "";
	$("foundBox").style.display = "none";
	$("selectedInfo").textContent = "Belum ada mahasiswa dipilih.";
	updateUndoBtn();
	updateStepper();
}

function persist() {
	try {
		localStorage.setItem(
			"nilai-mhs",
			JSON.stringify({
				fileName,
				fullAoa,
				headerRowIdx,
				metaRows,
				rawHeader,
				rawRows,
				components,
				compSeq,
				students,
				origSheetName,
				origCols,
				origMerges,
				origB64,
			}),
		);
	} catch (e) {}
	const ss = $("saveState");
	if (ss)
		ss.textContent = students.length
			? "● Tersimpan otomatis " + new Date().toLocaleTimeString("id-ID")
			: "";
}
(function restore() {
	try {
		const d = JSON.parse(localStorage.getItem("nilai-mhs") || "null");
		if (!d) return;
		fileName = d.fileName || "";
		fullAoa = d.fullAoa || [];
		headerRowIdx = d.headerRowIdx || 0;
		metaRows = d.metaRows || [];
		rawHeader = d.rawHeader || [];
		rawRows = d.rawRows || [];
		origSheetName = d.origSheetName || "";
		origCols = d.origCols || null;
		origMerges = d.origMerges || null;
		origB64 = d.origB64 || null;
		origBuffer = null;
		if (origB64) {
			try {
				origBuffer = b64ToBuf(origB64);
			} catch (_) {
				origBuffer = null;
			}
		}
		// kompatibilitas data lama (sebelum ada deteksi header): rekonstruksi fullAoa
		if (!fullAoa.length && rawHeader.length) {
			fullAoa = [rawHeader, ...rawRows];
			headerRowIdx = 0;
			metaRows = [];
		}
		// Migrasi format lama (uts/uas) ke komponen generik
		if (d.components && d.components.length) {
			components = d.components;
			compSeq = d.compSeq || d.components.length;
			students = (d.students || []).map((s) => ({
				name: s.name || "",
				nim: s.nim || "",
				grades: s.grades || {},
				row: s.row,
				r: s.r,
			}));
			// pastikan tiap mahasiswa punya semua kunci komponen
			students.forEach((s) => {
				components.forEach((c) => {
					if (!(c.id in s.grades)) s.grades[c.id] = null;
				});
			});
		} else if (
			d.students &&
			d.students.length &&
			("uts" in d.students[0] || "uas" in d.students[0])
		) {
			const cUts = d.students[0].cUts ?? -1;
			const cUas = d.students[0].cUas ?? -1;
			components = [
				{ id: "uts", label: "UTS", col: cUts },
				{ id: "uas", label: "UAS", col: cUas },
			];
			compSeq = 2;
			students = d.students.map((s) => ({
				name: s.name || "",
				nim: s.nim || "",
				grades: { uts: s.uts ?? null, uas: s.uas ?? null },
				row: s.row,
				r: s.r,
			}));
		} else {
			components = [];
			students = d.students || [];
		}
		// Normalisasi ke tepat 3 komponen (baris kosong bila kurang).
		while (components.length < 3) {
			components.push({ id: newCompId(), label: "", col: "" });
		}
		if (components.length > 3) {
			components = components.slice(0, 3);
			students.forEach((s) => {
				if (s.grades) {
					Object.keys(s.grades).forEach((k) => {
						if (!components.some((c) => c.id === k)) delete s.grades[k];
					});
				}
			});
		}
		students.forEach((s) => {
			if (!s.grades) s.grades = {};
			components.forEach((c) => {
				if (!(c.id in s.grades)) s.grades[c.id] = null;
			});
		});
		const mappingReady =
			components.length === 3 &&
			components.every(
				(c) => c.label && (c.col === -1 || parseInt(c.col) >= 0),
			);
		if (rawHeader.length) {
			buildHeaderRowSel();
			buildMappingUI();
			$("mapCard").style.display = "block";
			$("fileInfo").textContent =
				"File: " +
				(fileName || "(tanpa nama)") +
				" | " +
				rawRows.length +
				" baris.";
		}
		if (students.length && mappingReady) {
			rebuildDynamicUI();
			$("gradeCard").style.display = "block";
			$("tableCard").style.display = "block";
			$("fileInfo").textContent =
				"Data dipulihkan dari sesi sebelumnya (" +
				students.length +
				" mahasiswa, " +
				components.length +
				" komponen).";
			renderAll();
		}
	} catch (e) {}
})();
