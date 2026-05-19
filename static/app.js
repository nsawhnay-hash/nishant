'use strict';

// ── Character counter ──
const promptEl = document.getElementById('prompt');
const charCountEl = document.getElementById('charCount');
promptEl.addEventListener('input', () => {
  charCountEl.textContent = promptEl.value.length;
});

// ── CFG slider label ──
const cfgScale = document.getElementById('cfgScale');
const cfgValue = document.getElementById('cfgValue');
cfgScale.addEventListener('input', () => {
  cfgValue.textContent = parseFloat(cfgScale.value).toFixed(2);
});

// ── Panel switching ──
function showPanel(name) {
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('panel' + name.charAt(0).toUpperCase() + name.slice(1)).classList.add('active');
  document.getElementById('btn' + name.charAt(0).toUpperCase() + name.slice(1)).classList.add('active');

  if (name === 'history') loadHistory();
}

// ── Generation ──
let pollTimer = null;

async function submitGeneration() {
  const prompt = promptEl.value.trim();
  if (!prompt) {
    promptEl.focus();
    return;
  }

  const payload = {
    prompt,
    negative_prompt: document.getElementById('negPrompt').value.trim(),
    duration: document.getElementById('duration').value,
    aspect_ratio: document.getElementById('aspectRatio').value,
    cfg_scale: parseFloat(cfgScale.value),
  };

  setUIState('generating');

  try {
    const res = await fetch('/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload),
    });
    const { job_id, error } = await res.json();
    if (error) throw new Error(error);
    pollStatus(job_id);
  } catch (err) {
    showError(err.message || 'Failed to start generation.');
  }
}

function pollStatus(jobId) {
  let seen = 0;
  pollTimer = setInterval(async () => {
    try {
      const res = await fetch(`/status/${jobId}`);
      const job = await res.json();

      // Append new log lines
      const logBox = document.getElementById('logBox');
      const newLogs = job.logs.slice(seen);
      seen = job.logs.length;
      newLogs.forEach(msg => {
        const div = document.createElement('div');
        div.className = 'log-line' + (msg.startsWith('Done') ? ' done' : '');
        div.textContent = `› ${msg}`;
        logBox.appendChild(div);
        logBox.scrollTop = logBox.scrollHeight;
      });

      document.getElementById('progressStatus').textContent =
        job.status === 'queued' ? 'Queued…' : 'Generating…';

      if (job.status === 'completed') {
        clearInterval(pollTimer);
        showResult(job.video_url);
      } else if (job.status === 'failed') {
        clearInterval(pollTimer);
        showError(job.error || 'Generation failed.');
      }
    } catch {
      // network hiccup — keep polling
    }
  }, 1500);
}

// ── UI state helpers ──
function setUIState(state) {
  const submit = document.getElementById('btnSubmit');
  const progress = document.getElementById('progressCard');
  const result = document.getElementById('resultCard');
  const error = document.getElementById('errorCard');
  const logBox = document.getElementById('logBox');

  submit.disabled = state === 'generating';
  progress.style.display = state === 'generating' ? '' : 'none';
  result.style.display = 'none';
  error.style.display = 'none';

  if (state === 'generating') {
    logBox.innerHTML = '';
  }
}

function showResult(videoUrl) {
  document.getElementById('progressCard').style.display = 'none';
  document.getElementById('resultCard').style.display = '';
  document.getElementById('btnSubmit').disabled = false;

  const video = document.getElementById('resultVideo');
  const dl = document.getElementById('downloadBtn');
  video.src = videoUrl;
  video.play().catch(() => {});
  dl.href = videoUrl;
  dl.download = `video-${Date.now()}.mp4`;
}

function showError(msg) {
  document.getElementById('progressCard').style.display = 'none';
  document.getElementById('errorCard').style.display = '';
  document.getElementById('errorMsg').textContent = msg;
  document.getElementById('btnSubmit').disabled = false;
}

function resetForm() {
  clearInterval(pollTimer);
  setUIState('idle');
  document.getElementById('resultCard').style.display = 'none';
  document.getElementById('errorCard').style.display = 'none';
  document.getElementById('progressCard').style.display = 'none';
  document.getElementById('btnSubmit').disabled = false;
  promptEl.focus();
}

// ── History ──
async function loadHistory() {
  const grid = document.getElementById('historyGrid');
  try {
    const res = await fetch('/history');
    const items = await res.json();
    if (!items.length) {
      grid.innerHTML = '<p class="empty-state">No videos generated yet.</p>';
      return;
    }
    grid.innerHTML = items
      .map(
        item => `
      <div class="history-item">
        <video src="${item.video_url}" controls muted playsinline></video>
        <div class="history-item-footer">
          <a href="${item.video_url}" download="video.mp4" class="btn-download" style="font-size:12px;padding:7px 12px;">
            <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
            Download
          </a>
        </div>
      </div>`,
      )
      .join('');
  } catch {
    grid.innerHTML = '<p class="empty-state">Failed to load history.</p>';
  }
}
