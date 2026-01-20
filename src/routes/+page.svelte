<script>
  import { onMount } from 'svelte';

  /* ===============================
     KONSTANTA PROMPT AWAL
  =============================== */
  const dataawal = `Buatkan struktur notula berdasarkan dokumen hasil transkrip rapat
(Microsoft Teams / Zoom / Google Meet).

Struktur harus mencakup bagian-bagian utama dan sub-bagiannya, serta poin-poin penting,
dengan format berikut:

1. Judul Notula
2. Dasar
3. Waktu dan Tempat
4. Peserta
5. Agenda (ringkasan detail dari pembahasan)
6. Kesimpulan

Berikut adalah teks hasil transkrip:\n`;

  let prompt = dataawal;
  let output = '';
  let loading = false;

  /* ===============================
     HANDLE UPLOAD TXT
  =============================== */
  function handleFile(event) {
    const file = event.target.files[0];
    if (!file || file.type !== 'text/plain') {
      alert('Silakan upload file .txt');
      return;
    }

    const reader = new FileReader();
    reader.onload = e => {
      // 🔥 prompt awal tetap, isi file di-append
      prompt = dataawal + '\n' + e.target.result;
    };
    reader.readAsText(file, 'UTF-8');
  }

  /* ===============================
     GENERATE VIA BACKEND SVELTEKIT
  =============================== */
  async function generate() {
    if (!prompt.trim()) {
      alert('Prompt kosong');
      return;
    }

    loading = true;
    output = '';

    try {
      const res = await fetch('/api', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ prompt })
      });

      const data = await res.json();
      output = data.text || 'Tidak ada respon dari AI';
    } catch (err) {
      console.error(err);
      output = 'Terjadi kesalahan saat memproses data';
    } finally {
      loading = false;
    }
  }

  /* ===============================
     DOWNLOAD DOCX
  =============================== */
  async function downloadDocx() {
    if (!output.trim()) {
      alert('Tidak ada konten untuk diunduh');
      return;
    }

    const {
      Document,
      Packer,
      Paragraph,
      HeadingLevel,
      TextRun
    } = window.docx;

    // Parsing **bold**
    const parseBold = (line) => {
      const parts = [];
      const regex = /\*\*(.*?)\*\*/g;
      let lastIndex = 0;
      let match;

      while ((match = regex.exec(line)) !== null) {
        if (match.index > lastIndex) {
          parts.push(new TextRun(line.substring(lastIndex, match.index)));
        }
        parts.push(new TextRun({ text: match[1], bold: true }));
        lastIndex = regex.lastIndex;
      }

      if (lastIndex < line.length) {
        parts.push(new TextRun(line.substring(lastIndex)));
      }

      return parts;
    };

    const lines = output.split('\n');
    const paragraphs = [];

    lines.forEach((line, index) => {
      line = line.trim();
      if (!line) return;

      const cleaned = line.replace(/^\d+\.\s+/, '');

      // Heading pertama
      if (index === 0) {
        paragraphs.push(
          new Paragraph({
            heading: HeadingLevel.HEADING_1,
            spacing: { after: 300 },
            children: parseBold(cleaned)
          })
        );
        return;
      }

      // Numbering
      if (/^\d+\.\s+/.test(line)) {
        paragraphs.push(
          new Paragraph({
            numbering: { reference: 'num', level: 0 },
            spacing: { after: 150 },
            children: parseBold(cleaned)
          })
        );
        return;
      }

      // Paragraf biasa
      paragraphs.push(
        new Paragraph({
          spacing: { after: 200 },
          children: parseBold(cleaned)
        })
      );
    });

    const doc = new Document({
      numbering: {
        config: [
          {
            reference: 'num',
            levels: [
              {
                level: 0,
                format: 'decimal',
                text: '%1.',
                alignment: 'left'
              }
            ]
          }
        ]
      },
      sections: [
        {
          properties: {
            page: {
              margin: {
                top: 720,
                bottom: 720,
                left: 1000,
                right: 1000
              }
            }
          },
          children: paragraphs
        }
      ]
    });

    const blob = await Packer.toBlob(doc);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'notulAi_output.docx';
    a.click();
  }

  /* ===============================
     LOAD DOCX LIB (CLIENT ONLY)
  =============================== */
  onMount(() => {
    const script = document.createElement('script');
    script.src = 'https://cdn.jsdelivr.net/npm/docx@9.5.1/dist/index.iife.js';
    document.body.appendChild(script);
  });
</script>

<svelte:head>
  <title>NotulAI KPPN Lhokseumawe</title>
  <link
    href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css"
    rel="stylesheet"
  />
</svelte:head>

<div class="container mt-5">
  <h1 class="text-center mb-4">💡 NotulAI By KPPN Lhokseumawe</h1>

  <div class="card shadow p-4 mb-4">
    <div class="mb-3">
      <label class="form-label">Upload Transkrip (.txt)</label>
      <input
        type="file"
        class="form-control"
        accept=".txt"
        on:change={handleFile}
      />
    </div>

    <div class="mb-3">
      <label class="form-label">Prompt</label>
      <textarea
        class="form-control"
        rows="10"
        bind:value={prompt}
      ></textarea>
    </div>

    <button class="btn btn-primary w-100" on:click={generate}>
      🔮 Generate Notula
    </button>

    {#if loading}
      <div class="text-center mt-3">
        <div class="spinner-border text-primary"></div>
      </div>
    {/if}
  </div>

  {#if output}
    <div class="card shadow p-4">
      <h5>Hasil:</h5>
      <pre style="white-space: pre-wrap">{output}</pre>
    </div>

    <button class="btn btn-success mt-3" on:click={downloadDocx}>
      ⬇️ Download DOCX
    </button>
  {/if}
</div>
