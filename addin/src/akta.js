/* Logika Word untuk Aktara. Hanya memakai WordApi 1.1 agar jalan di Word 2016+ (Windows/Mac), 2019, 2021, 2024, 365 dan Word Online.
   Kode ES5 tanpa arrow function / async, karena Word 2016 MSI di Windows memakai engine IE11. */
(function () {
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
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document></pkg:xmlData></pkg:part>' +
    '<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"><pkg:xmlData>' +
    '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    '<w:style w:type="paragraph" w:customStyle="1" w:styleId="AktaGaris"><w:name w:val="' + STYLE + '"/><w:basedOn w:val="Normal"/><w:qFormat/>' +
    '<w:pPr><w:tabs><w:tab w:val="right" w:leader="hyphen" w:pos="9638"/></w:tabs><w:spacing w:after="0" w:line="360" w:lineRule="auto"/><w:jc w:val="left"/></w:pPr>' +
    '<w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:sz w:val="24"/></w:rPr></w:style>' +
    '</w:styles></pkg:xmlData></pkg:part></pkg:package>';

  function status(msg, err) {
    var el = document.getElementById("status");
    el.textContent = msg; el.className = err ? "status err" : "status ok";
  }
  function fail(e) { status("Gagal: " + (e && e.message ? e.message : e), true); }

  function pastikanStyle(ctx) {
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

  /* 1. Garis akta: pasang tab-leader di akhir setiap paragraf terpilih */
  function pasangGaris() {
    Word.run(function (ctx) {
      return pastikanStyle(ctx).then(function () {
        var ps = ctx.document.getSelection().paragraphs; ps.load("items/text");
        return ctx.sync().then(function () {
          var n = 0;
          ps.items.forEach(function (p) {
            p.style = STYLE;
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

  /* 3. Template akta dengan isian (content control ber-tag) */
  var TEMPLATE = {
    kuasa: { judul: "KUASA", isi: ["Berhadapan dengan saya, {NOTARIS}, Notaris di {KEDUDUKAN}, dengan dihadiri oleh saksi-saksi yang saya, Notaris, kenal dan akan disebut pada bagian akhir akta ini:",
      "Tuan/Nyonya {NAMA_PIHAK1}, lahir di {TTL_PIHAK1}, Warga Negara Indonesia, bertempat tinggal di {ALAMAT_PIHAK1}, pemegang Nomor Induk Kependudukan {NIK_PIHAK1};",
      "– selanjutnya disebut Pemberi Kuasa.",
      "Penghadap menerangkan dengan ini memberi kuasa kepada Tuan/Nyonya {NAMA_PIHAK2}, bertempat tinggal di {ALAMAT_PIHAK2}, pemegang Nomor Induk Kependudukan {NIK_PIHAK2};",
      "– selanjutnya disebut Penerima Kuasa.",
      "KHUSUS untuk {OBJEK_KUASA}."] },
    jualbeli: { judul: "PERJANJIAN PENGIKATAN JUAL BELI", isi: ["Berhadapan dengan saya, {NOTARIS}, Notaris di {KEDUDUKAN}, dengan dihadiri oleh saksi-saksi yang saya, Notaris, kenal dan akan disebut pada bagian akhir akta ini:",
      "I. Tuan/Nyonya {NAMA_PIHAK1}, lahir di {TTL_PIHAK1}, bertempat tinggal di {ALAMAT_PIHAK1}, pemegang Nomor Induk Kependudukan {NIK_PIHAK1};",
      "– selanjutnya disebut Pihak Pertama.",
      "II. Tuan/Nyonya {NAMA_PIHAK2}, lahir di {TTL_PIHAK2}, bertempat tinggal di {ALAMAT_PIHAK2}, pemegang Nomor Induk Kependudukan {NIK_PIHAK2};",
      "– selanjutnya disebut Pihak Kedua.",
      "Para penghadap menerangkan bahwa Pihak Pertama mengikatkan diri untuk menjual kepada Pihak Kedua {OBJEK} dengan harga {HARGA}."] }
  };
  var aktaTerakhir = null;
  function buatTemplate(kode, nomorAngka, sementara) {
    var t = TEMPLATE[kode], nomor = Terbilang.angka(nomorAngka);
    aktaTerakhir = { nomor: nomorAngka, sifat_akta: t.judul.charAt(0) + t.judul.slice(1).toLowerCase() };
    Word.run(function (ctx) {
      var body = ctx.document.getSelection();
      var awal = [t.judul, "Nomor : " + nomor, "Pada hari ini, {HARI}, tanggal {TANGGAL}, pukul {JAM}."].concat(t.isi);
      var last = body;
      // Sisip dari bawah ke atas lewat "After" berurutan
      var paras = [];
      awal.forEach(function (baris, i) {
        last = i === 0 ? body.insertParagraph(baris, "Before") : last.insertParagraph(baris, "After");
        paras.push(last);
      });
      return ctx.sync().then(function () {
        paras[0].alignment = "Centered"; paras[0].font.bold = true;
        paras[1].alignment = "Centered";
        // Bungkus setiap {FIELD} menjadi content control ber-tag
        var semua = [];
        paras.forEach(function (p) { var s = p.search("\\{[A-Z0-9_]@\\}", { matchWildcards: true }); s.load("items/text"); semua.push(s); });
        return ctx.sync().then(function () {
          semua.forEach(function (s) {
            s.items.forEach(function (r) {
              var tag = r.text.replace(/[{}]/g, "");
              var cc = r.insertContentControl(); cc.tag = tag; cc.title = tag.replace(/_/g, " "); cc.appearance = "Tags";
            });
          });
          return ctx.sync().then(function () { status("Template " + t.judul + " dibuat. Isi data di tab Pihak."); });
        });
      });
    }).catch(fail);
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
        var n = 0;
        list.forEach(function (x) { x[0].items.forEach(function (cc) { cc.insertText(x[1], "Replace"); n++; }); });
        return ctx.sync().then(function () { status(n + " isian diperbarui."); });
      });
    }).catch(fail);
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
      var p1 = document.querySelector('[data-tag="NAMA_PIHAK1"]').value, p2 = document.querySelector('[data-tag="NAMA_PIHAK2"]').value;
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
    document.getElementById("btnIsi").onclick = function () {
      var d = {}; var els = document.querySelectorAll("[data-tag]");
      for (var i = 0; i < els.length; i++) d[els[i].getAttribute("data-tag")] = els[i].value;
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
})();
