(function () {
  const canvas = document.getElementById('bg-canvas');
  const ctx    = canvas.getContext('2d');

  let W, H, nodes = [], mouse = { x: -999, y: -999 };
  const NODE_COUNT = 70;
  const CONNECT_DIST = 140;
  const COLORS = ['#c8873a', '#3b82f6', '#1e293b'];

  function resize() {
    W = canvas.width  = window.innerWidth;
    H = canvas.height = window.innerHeight;
  }

  function initNodes() {
    nodes = Array.from({ length: NODE_COUNT }, () => ({
      x: Math.random() * W,
      y: Math.random() * H,
      vx: (Math.random() - 0.5) * 0.35,
      vy: (Math.random() - 0.5) * 0.35,
      r: Math.random() * 1.8 + 0.8,
      color: COLORS[Math.floor(Math.random() * COLORS.length)],
    }));
  }

  function draw() {
    ctx.clearRect(0, 0, W, H);

    for (let i = 0; i < nodes.length; i++) {
      for (let j = i + 1; j < nodes.length; j++) {
        const a = nodes[i], b = nodes[j];
        const d = Math.hypot(a.x - b.x, a.y - b.y);
        if (d < CONNECT_DIST) {
          const alpha = (1 - d / CONNECT_DIST) * 0.18;
          ctx.beginPath();
          ctx.strokeStyle = `rgba(200,135,58,${alpha})`;
          ctx.lineWidth = 0.6;
          ctx.moveTo(a.x, a.y);
          ctx.lineTo(b.x, b.y);
          ctx.stroke();
        }
      }
    }

    nodes.forEach(n => {
      const dm = Math.hypot(n.x - mouse.x, n.y - mouse.y);
      if (dm < 160) {
        const f = (1 - dm / 160) * 0.012;
        n.vx += (mouse.x - n.x) * f;
        n.vy += (mouse.y - n.y) * f;
      }

      n.x += n.vx;
      n.y += n.vy;

      const speed = Math.hypot(n.vx, n.vy);
      if (speed > 0.9) { n.vx /= speed / 0.9; n.vy /= speed / 0.9; }

      if (n.x < 0 || n.x > W) n.vx *= -1;
      if (n.y < 0 || n.y > H) n.vy *= -1;

      ctx.beginPath();
      ctx.arc(n.x, n.y, n.r, 0, Math.PI * 2);
      ctx.fillStyle = n.color === '#c8873a'
        ? 'rgba(200,135,58,0.7)'
        : n.color === '#3b82f6'
          ? 'rgba(59,130,246,0.4)'
          : 'rgba(255,255,255,0.12)';
      ctx.fill();
    });

    requestAnimationFrame(draw);
  }

  window.addEventListener('resize', () => { resize(); initNodes(); });
  window.addEventListener('mousemove', e => { mouse.x = e.clientX; mouse.y = e.clientY; });

  resize();
  initNodes();
  draw();
})();