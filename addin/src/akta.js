/* Logika Word untuk Aktara. Hanya memakai WordApi 1.1 agar jalan di Word 2016+ (Windows/Mac), 2019, 2021, 2024, 365 dan Word Online.
   Kode ES5 tanpa arrow function / async, karena Word 2016 MSI di Windows memakai engine IE11. */
(function (root) {
  "use strict";
  var STYLE = "Akta Garis";
  // Paragraph style dengan tab kanan ber-leader strip di batas kanan teks: garis otomatis mengikuti panjang teks setelah diedit.
  var STYLE_OOXML =
    '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
    '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part>' +
    '<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships></pkg:xmlData></pkg:part>' +
    '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="AktaGaris"/></w:pPr></w:p></w:body></w:document></pkg:xmlData></pkg:part>' +
    '<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"><pkg:xmlData>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="AktaGaris"><w:name w:val="' + STYLE + '"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:tabs><w:tab w:val="right" w:leader="hyphen" w:pos="9638"/></w:tabs><w:spacing w:after="0" w:line="500" w:lineRule="exact"/><w:jc w:val="both"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="24"/></w:rPr></w:style>' +
    '</w:styles></pkg:xmlData></pkg:part></pkg:package>';

  function status(msg, err) {
    var el = document.getElementById("status");
    el.textContent = msg; el.className = err ? "status err" : "status ok";
  }
  function fail(e) {
    var info = "";
    try { if (e && e.debugInfo) info = " [" + (e.debugInfo.code || "") + " @ " + (e.debugInfo.errorLocation || "") + "]"; } catch (x) {}
    status("Gagal: " + (e && e.message ? e.message : e) + info, true);
  }

  function pastikanStyle(ctx) {
    // Paragraf dalam paket memakai style tersebut, sebab Word Online membuang style yang tidak dipakai.
    // Menyisipkan paket berisi definisi style; Word menambahkan style bila belum ada. Paragraf kosong sisaannya dihapus.
    var r = ctx.document.body.insertParagraph("", "End");
    r.insertOoxml(STYLE_OOXML, "Replace");
    return ctx.sync().then(function () {
      var ps = ctx.document.body.paragraphs; ps.load("items/text");
      return ctx.sync().then(function () {
        var last = ps.items[ps.items.length - 1];
        if (last.text === "") last.delete();
        return ctx.sync();
      });
    });
  }

  /* 1. Garis akta: tab-leader strip di akhir paragraf terpilih (judul rata tengah & paragraf kosong dilewati) */
  function pasangGaris() {
    Word.run(function (ctx) {
      return pastikanStyle(ctx).then(function () {
        var ps = ctx.document.getSelection().paragraphs; ps.load("items/text,items/alignment");
        return ctx.sync().then(function () {
          var n = 0;
          ps.items.forEach(function (p) {
            if (p.alignment === "Centered" || p.alignment === "Right" || !p.text) return;
            var rata = p.alignment;
            p.style = STYLE; p.alignment = rata;
            if (!/\t$/.test(p.text)) { p.insertText("\t", "End"); n++; }
          });
          return ctx.sync().then(function () { status("Garis dipasang pada " + n + " paragraf."); });
        });
      });
    }).catch(fail);
  }
  /* Hapus garis: buang tab di akhir paragraf terpilih */
  function hapusGaris() {
    Word.run(function (ctx) {
      var ps = ctx.document.getSelection().paragraphs; ps.load("items/text");
      return ctx.sync().then(function () {
        var hasil = [];
        ps.items.forEach(function (p) {
          if (/\t$/.test(p.text)) { var t = p.search("^t"); t.load("items"); hasil.push(t); }
        });
        return ctx.sync().then(function () {
          hasil.forEach(function (t) { if (t.items.length) t.items[t.items.length - 1].insertText("", "Replace"); });
          return ctx.sync().then(function () { status("Garis dihapus dari " + hasil.length + " paragraf."); });
        });
      });
    }).catch(fail);
  }

  /* 2. Sisip teks di kursor */
  function sisip(teks) {
    Word.run(function (ctx) {
      ctx.document.getSelection().insertText(teks, "Replace");
      return ctx.sync().then(function () { status("Disisipkan: " + teks); });
    }).catch(fail);
  }

  function label(tag) { return (root.AktaraLabel && AktaraLabel[tag]) || tag.replace(/_/g, " "); }
  function esc(t) { return String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/"/g, "&quot;"); }

  /* 3. Template akta (dari templates.js, sumber: templates/akta.json) dengan isian ber-tag */
  var aktaTerakhir = null;
  function buatTemplate(kode, nomorAngka) {
    var t = AktaraTemplates[kode];
    aktaTerakhir = { nomor: nomorAngka, sifat_akta: t.sifat };
    Word.run(function (ctx) {
      var sel = ctx.document.getSelection();
      var baris = [], tengah = [];
      t.judul.forEach(function (j) { baris.push(j); tengah.push("judul"); });
      baris.push("Nomor : " + Terbilang.nomorAkta(nomorAngka)); tengah.push("nomor");
      t.isi.forEach(function (x) {
        if (x.charAt(0) === "^") { baris.push(x.slice(1)); tengah.push("tengah"); }
        else { baris.push(x); tengah.push(""); }
      });
      // Setiap baris disusun per potongan: teks biasa disisipkan apa adanya, {ISIAN} langsung dijadikan
      // content control. Tanpa pencarian (search) - lebih andal di Word Online untuk paragraf panjang.
      var paras = [], last = sel;
      baris.forEach(function (b, i) {
        var potong = b.split(/(\{[A-Z0-9_]+\})/);
        last = i === 0 ? sel.insertParagraph(potong[0], "Before") : last.insertParagraph(potong[0], "After");
        // Seluruh teks paragraf disisipkan dulu, baru isian dibungkus content control. Bila content
        // control dibuat lebih awal, teks berikutnya yang disisipkan di "End" ikut masuk ke dalamnya
        // (terjadi pada paragraf yang diawali isian di Word Online).
        var isian = [];
        for (var k = 1; k < potong.length; k++) {
          if (!potong[k]) continue;
          var r = last.insertText(potong[k], "End");
          if (k % 2 === 1) isian.push([r, potong[k].slice(1, -1)]);
        }
        isian.forEach(function (x) { var cc = x[0].insertContentControl(); cc.tag = x[1]; cc.title = label(x[1]); });
        paras.push(last);
      });
      // Paragraf yang diawali {ISIAN} membuat proxy paragrafnya tidak andal di Word Online, sehingga
      // format diterapkan lewat koleksi paragraf dari seluruh rentang akta.
      var rentang = paras[0].getRange("Whole").expandTo(paras[paras.length - 1].getRange("Whole"));
      var semua = rentang.paragraphs; semua.load("items");
      return ctx.sync().then(function () {
        if (semua.items.length === paras.length) paras = semua.items;
        rentang.font.name = "Courier New"; rentang.font.size = 12;
        paras.forEach(function (p, i) {
          // Format praktik akta: Courier New 12, spasi tepat 25 pt, rata kiri-kanan
          p.font.name = "Courier New"; p.font.size = 12; p.lineSpacing = 25; p.spaceAfter = 0;
          p.alignment = tengah[i] ? "Centered" : "Justified";
          if (tengah[i] === "judul") p.font.bold = true;
        });
        return ctx.sync();
      }).then(function () { status("Akta " + t.sifat + " nomor " + nomorAngka + " dibuat. Buka tab Pihak > Muat isian."); });
    }).catch(fail);
  }

  /* Muat daftar isian unik dari dokumen ke tab Pihak (label lengkap, NIK divalidasi) */
  function muatIsian() {
    Word.run(function (ctx) {
      var cc = ctx.document.contentControls; cc.load("items/tag,items/title,items/text");
      return ctx.sync().then(function () {
        var wadah = document.getElementById("daftarIsian"), sudah = {}, html = "", n = 0;
        cc.items.forEach(function (c) {
          if (!c.tag || sudah[c.tag]) return; sudah[c.tag] = 1; n++;
          var isi = /^\{.*\}$/.test(c.text) ? "" : c.text;
          html += '<label>' + esc(label(c.tag)) + '</label><input data-tag="' + esc(c.tag) + '" value="' + esc(isi) + '">';
          if (/^NIK_/.test(c.tag)) html += '<div class="peringatan" data-cek="' + esc(c.tag) + '"></div>';
        });
        wadah.innerHTML = html || '<p class="hint">Belum ada isian. Buat akta dari template dulu.</p>';
        var niks = wadah.querySelectorAll('input[data-tag^="NIK_"]');
        for (var i = 0; i < niks.length; i++) { niks[i].oninput = periksaNIK; periksaNIK.call(niks[i]); }
        status(n + " isian dimuat.");
      });
    }).catch(fail);
  }
  function periksaNIK() {
    var tag = this.getAttribute("data-tag"), el = document.querySelector('[data-cek="' + tag + '"]');
    if (!el) return;
    if (!this.value) { el.textContent = ""; this.className = ""; return; }
    var pesan = AktaraValidasi.cekNIK(this.value), info = AktaraValidasi.infoNIK(this.value);
    this.className = pesan ? "salah" : "";
    el.className = pesan ? "peringatan" : "info";
    el.textContent = pesan || ("Sah. Tanggal lahir menurut NIK: " + info.tanggal + " (" + (info.perempuan ? "perempuan" : "laki-laki") + ").");
  }

  /* 4. Isi data pihak ke seluruh content control dengan tag yang sama */
  function isiField(data) {
    Word.run(function (ctx) {
      var list = [];
      Object.keys(data).forEach(function (tag) {
        if (!data[tag]) return;
        var c = ctx.document.contentControls.getByTag(tag); c.load("items"); list.push([c, data[tag]]);
      });
      return ctx.sync().then(function () {
        // Diisi satu per satu (sync per content control): Word Online kadang menolak banyak Replace
        // sekaligus (GeneralException). Bila Replace gagal, isi dikosongkan lalu disisipkan ulang.
        var tugas = [], n = 0, gagal = 0;
        list.forEach(function (x) { x[0].items.forEach(function (cc) { tugas.push([cc, x[1]]); }); });
        function langkah(i) {
          if (i >= tugas.length) return null;
          var cc = tugas[i][0], nilai = tugas[i][1];
          // Teks pengganti mengikuti format akta (Word Online memakai font bawaan saat Replace)
          cc.insertText(nilai, "Replace"); cc.font.name = "Courier New"; cc.font.size = 12;
          return ctx.sync().then(function () { n++; }, function () {
            cc.clear(); cc.insertText(nilai, "Start"); cc.font.name = "Courier New"; cc.font.size = 12;
            return ctx.sync().then(function () { n++; }, function () { gagal++; });
          }).then(function () { return langkah(i + 1); });
        }
        return (window.Promise || OfficeExtension.Promise).resolve(langkah(0)).then(function () {
          status(n + " isian diperbarui." + (gagal ? " " + gagal + " gagal diisi." : ""), gagal > 0);
        });
      });
    }).catch(fail);
  }

  /* Terbilangkan angka terpilih: "2026" -> "2026 (dua ribu dua puluh enam)" */
  function terbilangkanPilihan() {
    Word.run(function (ctx) {
      var sel = ctx.document.getSelection(); sel.load("text");
      return ctx.sync().then(function () {
        var t = (sel.text || "").replace(/\s+$/, ""), angka = t.replace(/[^\d,]/g, "");
        if (!/\d/.test(angka)) return status("Blok angka di dokumen terlebih dahulu.", true);
        var huruf = angka.indexOf(",") >= 0 ? Terbilang.desimal(angka) : Terbilang.terbilang(parseInt(angka.replace(/,/g, ""), 10));
        sel.insertText(t + " (" + huruf + ")", "Replace");
        return ctx.sync().then(function () { status("Diterbilangkan: " + huruf); });
      });
    }).catch(fail);
  }

  /* Periksa akta sebelum cetak */
  function periksaAkta() {
    Word.run(function (ctx) {
      var cc = ctx.document.contentControls; cc.load("items/tag,items/text");
      var ps = ctx.document.body.paragraphs; ps.load("items/text,items/alignment");
      return ctx.sync().then(function () {
        var kosong = {}, tanpaGaris = 0, adaNomor = false;
        cc.items.forEach(function (c) { if (!c.text || /^\{.*\}$/.test(c.text)) kosong[c.tag] = 1; });
        ps.items.forEach(function (p) {
          if (/^Nomor\s*:\s*\d+/.test(p.text)) adaNomor = true;
          if (p.text && p.alignment !== "Centered" && p.alignment !== "Right" && !/\t$/.test(p.text)) tanpaGaris++;
        });
        var daftar = Object.keys(kosong), html = "<ul class='cek'>";
        html += adaNomor ? "<li class='info'>Nomor akta tercantum.</li>" : "<li class='peringatan'>Nomor akta belum tercantum.</li>";
        if (daftar.length) { html += "<li class='peringatan'>" + daftar.length + " isian masih kosong:<ul>";
          daftar.forEach(function (t) { html += "<li>" + esc(label(t)) + "</li>"; }); html += "</ul></li>"; }
        else html += "<li class='info'>Semua isian sudah terisi.</li>";
        html += tanpaGaris ? "<li class='peringatan'>" + tanpaGaris + " paragraf belum diberi garis akta.</li>" : "<li class='info'>Semua paragraf sudah bergaris.</li>";
        document.getElementById("hasilCek").innerHTML = html + "</ul>";
        status(daftar.length || tanpaGaris || !adaNomor ? "Akta belum siap cetak." : "Akta siap dicetak.", !!(daftar.length || tanpaGaris || !adaNomor));
      });
    }).catch(fail);
  }

  /* Tampilkan repertorium bulan berjalan */
  function lihatRepertorium() {
    var d = new Date();
    AktaraServer.daftarRepertorium(d.getFullYear(), d.getMonth() + 1, function (err, r) {
      var w = document.getElementById("daftarRep");
      if (err) { w.innerHTML = ""; return status("Gagal memuat repertorium: " + err.pesan, true); }
      var h = "<table class='rep'><tr><th>No.</th><th>Tanggal</th><th>Sifat akta</th><th>Penghadap</th></tr>";
      r.data.forEach(function (b) { h += "<tr><td>" + b.nomor + "</td><td>" + esc(b.tanggal) + "</td><td>" + esc(b.sifat_akta) + "</td><td>" + esc(b.penghadap) + "</td></tr>"; });
      w.innerHTML = r.data.length ? h + "</table>" : "<p class='hint'>Belum ada akta tercatat bulan ini.</p>";
      status(r.jumlah + " akta tercatat bulan ini.");
    });
  }

  /* 5. Minuta -> Salinan */
  function jadikanSalinan(notaris) {
    Word.run(function (ctx) {
      var b = ctx.document.body;
      var h = b.insertParagraph("SALINAN", "Start"); h.alignment = "Right"; h.font.bold = true;
      b.insertParagraph("– Diberikan sebagai SALINAN yang sama bunyinya.", "End");
      b.insertParagraph("Notaris di " + (notaris || "{KEDUDUKAN}"), "End");
      var c = b.insertParagraph("", "End");
      c.insertText("(tanda tangan dan cap jabatan)", "Start");
      return ctx.sync().then(function () { status("Dokumen diubah menjadi SALINAN."); });
    }).catch(fail);
  }

  /* 6. Lisensi & repertorium lewat AktaraServer (server.js) */
  function tampilLisensi(st) {
    var el = document.getElementById("infoLisensi");
    var teks = { belum: "Belum diaktifkan.", aktif: "Aktif — " + (st.kantor || ""),
      offline: "Offline — lisensi tetap berlaku " + st.sisaHariOffline + " hari lagi.",
      kedaluwarsa: "Lisensi perlu dicek ulang. Sambungkan internet atau hubungi dukungan Aktara." };
    el.textContent = (st.pesan ? st.pesan + " " : "") + teks[st.status];
    document.getElementById("formAktivasi").style.display = st.status === "belum" ? "block" : "none";
  }
  function bolehPakai() {
    var st = AktaraServer.statusLisensi().status;
    if (st === "aktif" || st === "offline") return true;
    status("Aktifkan lisensi Aktara di tab Lisensi terlebih dahulu.", true); return false;
  }

  function val(id) { return document.getElementById(id).value; }

  Office.onReady(function () {
    document.getElementById("btnGaris").onclick = pasangGaris;
    document.getElementById("btnHapusGaris").onclick = hapusGaris;
    document.getElementById("btnTanggal").onclick = function () {
      var p = val("inTanggal").split("-"); if (p.length !== 3) return status("Pilih tanggal dulu.", true);
      sisip(Terbilang.tanggal(new Date(+p[0], +p[1] - 1, +p[2])));
    };
    document.getElementById("btnJam").onclick = function () {
      var p = val("inJam").split(":"); if (p.length < 2) return status("Isi jam dulu.", true);
      sisip(Terbilang.jam(+p[0], +p[1], val("inZona")));
    };
    document.getElementById("btnRupiah").onclick = function () {
      var n = parseInt(val("inRupiah").replace(/\D/g, ""), 10); if (isNaN(n)) return status("Isi nominal dulu.", true);
      sisip(Terbilang.rupiah(n));
    };
    document.getElementById("btnAngka").onclick = function () {
      var n = parseInt(val("inRupiah").replace(/\D/g, ""), 10); if (isNaN(n)) return status("Isi angka dulu.", true);
      sisip(Terbilang.angka(n));
    };
    document.getElementById("btnTemplate").onclick = function () {
      if (!bolehPakai()) return;
      var kode = val("inTemplate");
      AktaraServer.nomorAkta(function (err, r) {
        buatTemplate(kode, r.nomor);
        if (r.sementara) setTimeout(function () { status("Offline: nomor " + r.nomor + " SEMENTARA. Cek ulang sebelum akta ditandatangani.", true); }, 1500);
      });
    };
    document.getElementById("btnRepertorium").onclick = function () {
      if (!aktaTerakhir) return status("Buat akta dari template dulu.", true);
      var ambil = function (t) { var e = document.querySelector('[data-tag="' + t + '"]'); return e ? e.value : ""; };
      var p1 = ambil("NAMA_PIHAK1"), p2 = ambil("NAMA_PIHAK2");
      var tgl = new Date(), iso = tgl.getFullYear() + "-" + ("0" + (tgl.getMonth() + 1)).slice(-2) + "-" + ("0" + tgl.getDate()).slice(-2);
      AktaraServer.catatRepertorium({ nomor: aktaTerakhir.nomor, tanggal: iso, sifat_akta: aktaTerakhir.sifat_akta,
        penghadap: [p1, p2].filter(function (x) { return x; }).join("; ") || "-" }, function (err) {
        status(err ? "Gagal mencatat repertorium: " + err.pesan : "Akta nomor " + aktaTerakhir.nomor + " tercatat di repertorium.", !!err);
      });
    };
    document.getElementById("btnAktivasi").onclick = function () {
      if (val("inServer")) AktaraServer.aturServer(val("inServer"));
      AktaraServer.aktivasi(val("inKode"), val("inNamaPc"), function (err, d) {
        if (err) return status("Aktivasi gagal: " + err.pesan, true);
        status("Aktivasi berhasil untuk " + d.kantor + "."); tampilLisensi(AktaraServer.statusLisensi());
      });
    };
    document.getElementById("btnMuat").onclick = muatIsian;
    document.getElementById("btnPilihan").onclick = terbilangkanPilihan;
    document.getElementById("btnCek").onclick = periksaAkta;
    document.getElementById("btnLihatRep").onclick = lihatRepertorium;
    document.getElementById("btnPersen").onclick = function () {
      var v = val("inRupiah").replace(/[^\d,.]/g, ""); if (!v) return status("Isi angka persen dulu.", true);
      sisip(Terbilang.persen(v));
    };
    var pilih = document.getElementById("inTemplate");
    Object.keys(AktaraTemplates).forEach(function (k) {
      var o = document.createElement("option"); o.value = k; o.text = AktaraTemplates[k].nama; pilih.appendChild(o);
    });
    document.getElementById("btnLuas").onclick = function () {
      var n = parseInt(val("inRupiah").replace(/\D/g, ""), 10); if (isNaN(n)) return status("Isi angka luas dulu.", true);
      sisip(Terbilang.luas(n));
    };
    document.getElementById("btnIsi").onclick = function () {
      var d = {}; var els = document.querySelectorAll("[data-tag]");
      var salah = [];
      for (var i = 0; i < els.length; i++) {
        var tg = els[i].getAttribute("data-tag"); d[tg] = els[i].value;
        if (/^NIK_/.test(tg) && els[i].value && AktaraValidasi.cekNIK(els[i].value)) salah.push(label(tg));
      }
      if (salah.length) status("Periksa kembali: " + salah.join(", ") + ". Data tetap diisikan.", true);
      isiField(d);
    };
    document.getElementById("btnSalinan").onclick = function () { jadikanSalinan(val("inKedudukan")); };
    var tabs = document.querySelectorAll(".tab");
    for (var i = 0; i < tabs.length; i++) tabs[i].onclick = function () {
      for (var j = 0; j < tabs.length; j++) { tabs[j].className = "tab"; document.getElementById(tabs[j].getAttribute("data-p")).style.display = "none"; }
      this.className = "tab on"; document.getElementById(this.getAttribute("data-p")).style.display = "block";
    };
    document.getElementById("inServer").value = AktaraServer.urlServer();
    tampilLisensi(AktaraServer.statusLisensi());
    AktaraServer.cekLisensi(function (e, st) { tampilLisensi(st); });
    status("Siap.");
  });
})(this);
