function getPrintHtml(filteredDataJson) {
  let data = [];
  try { data = JSON.parse(filteredDataJson); } catch (e) { return "Ralat data"; }
  
  const today = new Date().toLocaleDateString('ms-MY', { day: '2-digit', month: '2-digit', year: 'numeric' });
  const time = new Date().toLocaleTimeString('ms-MY', { hour: '2-digit', minute: '2-digit' });

  // 1. KIRAAN RINGKASAN KEWANGAN (Aggregates)
  const totalRekod = data.length;
  const totalKos = data.reduce((acc, r) => acc + (r.KOS_AKHIR || 0), 0);
  const totalBayar = data.reduce((acc, r) => acc + (r.JUMLAH_BAYARAN || 0), 0);
  const bakiBayar = totalKos - totalBayar;
  const peratusBayar = totalKos > 0 ? (totalBayar / totalKos) * 100 : 0;

  // Helper Format Duit
  const fmt = (num) => 'RM ' + (num || 0).toLocaleString('en-MY', { minimumFractionDigits: 2, maximumFractionDigits: 2 });

  // CSS & HTML STRUCTURE
  let html = `
    <html>
      <head>
        <title>Laporan Projek BPM</title>
        <style>
          @page { size: landscape; margin: 10mm; }
          body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 8pt; color: #333; -webkit-print-color-adjust: exact; }
          
          /* HEADER */
          .header-box { border-bottom: 2px solid #2563eb; padding-bottom: 10px; margin-bottom: 15px; display: flex; justify-content: space-between; align-items: flex-end; }
          h1 { font-size: 16pt; margin: 0; color: #1e293b; text-transform: uppercase; font-weight: 900; }
          .sub-title { font-size: 9pt; color: #64748b; margin-top: 2px; text-transform: uppercase; letter-spacing: 1px; }
          .meta { font-size: 8pt; text-align: right; color: #475569; font-weight: bold; }

          /* SUMMARY CARDS (ATAS) */
          .summary-container { display: flex; gap: 10px; margin-bottom: 20px; background: #f8fafc; padding: 10px; border: 1px solid #e2e8f0; border-radius: 6px; }
          .sum-card { flex: 1; padding: 8px 12px; background: #fff; border: 1px solid #cbd5e1; border-radius: 4px; border-left: 4px solid #64748b; }
          
          /* Warna Border Card */
          .sum-card.blue { border-left-color: #2563eb; }   /* Nilai Kontrak */
          .sum-card.green { border-left-color: #16a34a; }  /* Bayaran */
          .sum-card.orange { border-left-color: #f97316; } /* Baki */

          .sum-label { font-size: 7pt; color: #64748b; text-transform: uppercase; font-weight: bold; display: block; margin-bottom: 2px; }
          .sum-val { font-size: 11pt; font-weight: 800; color: #0f172a; font-family: 'Consolas', monospace; }
          .sum-sub { font-size: 7pt; color: #94a3b8; font-weight: 600; margin-top: 2px; }

          /* TABLE STYLES */
          table { width: 100%; border-collapse: collapse; border-spacing: 0; }
          th { background-color: #f1f5f9; color: #334155; font-weight: 800; text-transform: uppercase; font-size: 7pt; padding: 8px; border: 1px solid #cbd5e1; letter-spacing: 0.5px; }
          td { border: 1px solid #e2e8f0; padding: 8px; vertical-align: top; }
          
          /* COLUMN WIDTHS */
          .col-projek { width: 28%; }
          .col-pihak { width: 16%; }
          .col-masa { width: 18%; }
          .col-bayar { width: 14%; }
          .col-kos { width: 14%; }
          .col-status { width: 10%; }

          /* HELPER CLASSES */
          .label { font-size: 6pt; color: #94a3b8; text-transform: uppercase; font-weight: bold; display: block; margin-bottom: 1px; }
          .val { font-weight: 600; color: #0f172a; display: block; margin-bottom: 4px; }
          .badge { display: inline-block; padding: 2px 4px; border-radius: 3px; font-size: 6pt; font-weight: bold; border: 1px solid #ccc; margin-right: 3px; }
          .money { font-family: 'Consolas', monospace; text-align: right; font-weight: 700; }
          .section-group { border-bottom: 1px dashed #e2e8f0; padding-bottom: 4px; margin-bottom: 4px; }
          .section-group:last-child { border-bottom: none; margin-bottom: 0; }
          .stat-box { text-align: center; padding: 4px; border-radius: 4px; font-weight: bold; font-size: 7pt; border: 1px solid #ddd; }

          /* TOTAL ROW (BAWAH) */
          tr.total-row td { background-color: #f8fafc; border-top: 2px solid #64748b; font-weight: bold; color: #0f172a; padding-top: 8px; padding-bottom: 8px; }
        </style>
      </head>
      <body onload="window.print()">
        
        <div class="header-box">
          <div>
            <h1>Laporan Status Projek BPM</h1>
            <div class="sub-title">Senarai Terperinci Projek & Kontrak</div>
          </div>
          <div class="meta">
            Dicetak pada: ${today} (${time})
          </div>
        </div>

        <div class="summary-container">
           <div class="sum-card">
              <span class="sum-label">Bilangan Projek</span>
              <span class="sum-val">${totalRekod}</span>
              <div class="sum-sub">Rekod Terpilih</div>
           </div>
           <div class="sum-card blue">
              <span class="sum-label">Jumlah Nilai Kontrak</span>
              <span class="sum-val">${fmt(totalKos)}</span>
              <div class="sum-sub">Komitmen Keseluruhan</div>
           </div>
           <div class="sum-card green">
              <span class="sum-label">Jumlah Telah Dibayar</span>
              <span class="sum-val">${fmt(totalBayar)}</span>
              <div class="sum-sub">Prestasi: ${peratusBayar.toFixed(1)}%</div>
           </div>
           <div class="sum-card orange">
              <span class="sum-label">Baki Belum Bayar</span>
              <span class="sum-val">${fmt(bakiBayar)}</span>
              <div class="sum-sub">Tunggakan / Baki Kontrak</div>
           </div>
        </div>

        <table>
          <thead>
            <tr>
              <th class="col-projek">Maklumat Projek</th>
              <th class="col-pihak">Pihak Terlibat</th>
              <th class="col-masa">Garis Masa</th>
              <th class="col-bayar">Prestasi Kewangan</th>
              <th class="col-kos">Nilai Kontrak (RM)</th>
              <th class="col-status">Status</th>
            </tr>
          </thead>
          <tbody>
  `;

  data.forEach((r) => {
    // Format Nombor
    const kosAkhir = (r.KOS_AKHIR||0).toLocaleString('en-MY', {minimumFractionDigits:2});
    const bayarJum = (r.JUMLAH_BAYARAN||0).toLocaleString('en-MY', {minimumFractionDigits:2});
    const kosBase = (r.KOS_PECAHAN?.kontrak||0).toLocaleString('en-MY', {minimumFractionDigits:2});
    const kosSST = (r.KOS_PECAHAN?.sst||0).toLocaleString('en-MY', {minimumFractionDigits:2});

    // Warna Status
    let statusBg = '#f1f5f9'; let statusColor = '#334155';
    if(r.LOGIK_SIAP.includes('Siap')) { statusBg = '#dcfce7'; statusColor = '#166534'; } 
    else if(r.LOGIK_SIAP.includes('Lewat')) { statusBg = '#fee2e2'; statusColor = '#991b1b'; } 
    else if(r.LOGIK_SIAP.includes('Jadual')) { statusBg = '#dbeafe'; statusColor = '#1e40af'; }

    // HTML Baris Jadual
    html += `
      <tr>
        <td>
          <div class="val" style="font-size:9pt; line-height:1.3;">${r.NAMA_PAPARAN}</div>
          <div style="margin-top:4px;">
            <span class="badge" style="background:#f1f5f9;">${r['NO SEBUT HARGA']||'N/A'}</span>
            <span class="badge" style="background:#eff6ff; color:#1d4ed8;">${r['KATEGORI PERKHIDMATAN']||'-'}</span>
          </div>
          ${r.VOT_NAMA && r.VOT_NAMA !== '-' ? `<div style="margin-top:4px; font-size:7pt; color:#6b21a8;"><strong>VOT:</strong> ${r.VOT_NAMA}</div>` : ''}
        </td>
        <td>
          <div class="section-group">
            <span class="label">Kontraktor</span><span class="val">${r.SYARIKAT_NAMA}</span>
          </div>
          <div class="section-group">
            <span class="label">Penyelia</span><span class="val" style="font-size:7pt;">${r.PEGAWAI_NAMA}</span>
            <span style="font-size:7pt; color:#64748b;">${r.BAHAGIAN_NAMA}</span>
          </div>
        </td>
        <td>
          <div style="display:flex; justify-content:space-between; margin-bottom:2px;">
            <span class="label">Mula:</span> <strong>${r['TARIKH MULA']||'-'}</strong>
          </div>
          <div style="display:flex; justify-content:space-between; margin-bottom:4px;">
            <span class="label">Siap:</span> <strong>${r['TARIKH SIAP']||'-'}</strong>
          </div>
          ${r['EOT'] ? `<div style="background:#fff7ed; color:#c2410c; padding:2px 4px; border-radius:3px; font-size:7pt; margin-bottom:4px; border:1px solid #ffedd5;"><strong>EOT:</strong> ${r['EOT']}</div>` : ''}
          ${r.TEMPOH_DATA && r.TEMPOH_DATA.length > 0 ? 
            `<div style="margin-top:4px; border-top:1px dotted #ccc; padding-top:2px;">${r.TEMPOH_DATA.map(t => `<div style="display:flex; justify-content:space-between; font-size:7pt; color:#334155;"><span>${t.label}:</span> <span>${t.val}</span></div>`).join('')}</div>` 
            : ''}
        </td>
        <td>
          <div class="section-group">
            <span class="label">Progress Bayaran</span><span class="val" style="font-size:10pt;">${r.PROGRESS_MASA}%</span>
          </div>
          <div class="section-group">
            <span class="label">Jumlah Dibayar</span><span class="money" style="display:block;">${bayarJum}</span>
          </div>
        </td>
        <td>
          <div class="section-group">
            <span class="label">Kos Akhir</span><span class="money" style="display:block; font-size:9pt;">${kosAkhir}</span>
          </div>
          <div style="font-size:7pt; color:#64748b; text-align:right;">
            <div>Base: ${kosBase}</div>
            <div>Tax: ${kosSST}</div>
          </div>
        </td>
        <td style="vertical-align:middle;">
          <div class="stat-box" style="background:${statusBg}; color:${statusColor};">${r.LOGIK_SIAP}</div>
        </td>
      </tr>
    `;
  });

  // 3. BARIS JUMLAH (TOTAL FOOTER)
  html += `
          <tr class="total-row">
            <td colspan="3" style="text-align:right; padding-right:15px; font-size:9pt;">JUMLAH KESELURUHAN:</td>
            <td class="money">${fmt(totalBayar)}</td>
            <td class="money">${fmt(totalKos)}</td>
            <td style="font-size:7pt; text-align:center;">
               ${peratusBayar.toFixed(1)}% SIAP
            </td>
          </tr>
        </tbody>
      </table>
      
      <div style="margin-top:30px; font-size:7pt; color:#94a3b8; text-align:center;">
        Laporan dijana secara automatik oleh Sistem Dashboard BPM LKIM | Mukasurat Tamat
      </div>
    </body></html>`;
  
  return html;
}