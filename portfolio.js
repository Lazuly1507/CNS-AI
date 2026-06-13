(() => {
  const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)").matches;
  const finePointer = window.matchMedia("(pointer: fine)").matches;
  const transition = document.querySelector(".page-transition");
  const header = document.querySelector(".site-header");
  const nav = document.querySelector(".site-nav");
  const navToggle = document.querySelector(".nav-toggle");
  const progress = document.querySelector(".scroll-progress span");
  const backTop = document.querySelector(".back-top");

  window.setTimeout(() => transition?.classList.add("ready"), reducedMotion ? 0 : 650);

  function updateScroll() {
    const scrollable = document.documentElement.scrollHeight - window.innerHeight;
    const percent = scrollable > 0 ? (window.scrollY / scrollable) * 100 : 0;
    if (progress) progress.style.width = `${percent}%`;
    header?.classList.toggle("scrolled", window.scrollY > 28);
    backTop?.classList.toggle("show", window.scrollY > 600);
  }

  updateScroll();
  window.addEventListener("scroll", updateScroll, { passive: true });
  backTop?.addEventListener("click", () => window.scrollTo({ top: 0, behavior: reducedMotion ? "auto" : "smooth" }));

  navToggle?.addEventListener("click", () => {
    const open = nav?.classList.toggle("open") ?? false;
    navToggle.setAttribute("aria-expanded", String(open));
    document.body.classList.toggle("menu-open", open);
  });

  nav?.addEventListener("click", (event) => {
    if (!(event.target instanceof HTMLAnchorElement)) return;
    nav.classList.remove("open");
    navToggle?.setAttribute("aria-expanded", "false");
    document.body.classList.remove("menu-open");
  });

  if ("IntersectionObserver" in window && !reducedMotion) {
    const observer = new IntersectionObserver((entries) => {
      entries.forEach((entry) => {
        if (!entry.isIntersecting) return;
        entry.target.classList.add("visible");
        observer.unobserve(entry.target);
      });
    }, { threshold: 0.12, rootMargin: "0px 0px -35px" });
    document.querySelectorAll(".reveal").forEach((item) => observer.observe(item));
  } else {
    document.querySelectorAll(".reveal").forEach((item) => item.classList.add("visible"));
  }
  window.setTimeout(() => {
    document.querySelectorAll(".reveal").forEach((item) => item.classList.add("visible"));
    transition?.classList.add("ready");
  }, 1400);

  const lightbox = document.querySelector(".lightbox");
  document.querySelectorAll(".gallery-item[data-src], .tutorial-image[data-src], .context-image[data-src]").forEach((item) => {
    item.addEventListener("click", () => {
      if (!(lightbox instanceof HTMLDialogElement)) return;
      const image = lightbox.querySelector("img");
      const caption = lightbox.querySelector("p");
      if (image) {
        image.src = item.dataset.src ?? "";
        image.alt = item.dataset.caption ?? "";
      }
      if (caption) caption.textContent = item.dataset.caption ?? "";
      lightbox.showModal();
    });
  });
  lightbox?.querySelector(".lightbox-close")?.addEventListener("click", () => lightbox.close());
  lightbox?.addEventListener("click", (event) => {
    if (event.target === lightbox) lightbox.close();
  });

  document.querySelectorAll("a[href]").forEach((link) => {
    link.addEventListener("click", (event) => {
      const url = new URL(link.href, window.location.href);
      const samePageHash = url.pathname === window.location.pathname && url.hash;
      if (
        event.defaultPrevented ||
        link.target === "_blank" ||
        url.origin !== window.location.origin ||
        samePageHash ||
        reducedMotion
      ) return;
      event.preventDefault();
      transition?.classList.remove("ready");
      transition?.classList.add("leaving");
      window.setTimeout(() => { window.location.href = url.href; }, 620);
    });
  });

  document.querySelectorAll(".counter[data-target]").forEach((counter) => {
    const raw = counter.dataset.target ?? "0";
    const parts = raw.match(/^(\d+)(.*)$/);
    if (!parts) return;
    const target = Number.parseInt(parts[1], 10);
    if (!Number.isFinite(target) || reducedMotion) return;
    const suffix = parts[2];
    const width = parts[1].length;
    const padded = parts[1].startsWith("0");
    let started = false;
    const animate = () => {
      if (started) return;
      started = true;
      const start = performance.now();
      const step = (time) => {
        const ratio = Math.min((time - start) / 1000, 1);
        const value = Math.round(target * (1 - Math.pow(1 - ratio, 3)));
        counter.textContent = `${padded ? String(value).padStart(width, "0") : value}${suffix}`;
        if (ratio < 1) requestAnimationFrame(step);
      };
      requestAnimationFrame(step);
    };
    const observer = new IntersectionObserver((entries) => {
      if (entries[0].isIntersecting) {
        animate();
        observer.disconnect();
      }
    });
    observer.observe(counter);
  });

  if (finePointer && !reducedMotion) {
    const glow = document.querySelector(".cursor-glow");
    window.addEventListener("pointermove", (event) => {
      if (!(glow instanceof HTMLElement)) return;
      glow.style.opacity = "1";
      glow.style.left = `${event.clientX}px`;
      glow.style.top = `${event.clientY}px`;
    }, { passive: true });

    document.querySelectorAll(".tilt").forEach((item) => {
      item.addEventListener("pointermove", (event) => {
        const rect = item.getBoundingClientRect();
        const x = ((event.clientX - rect.left) / rect.width - .5) * 8;
        const y = ((event.clientY - rect.top) / rect.height - .5) * -8;
        item.style.transform = `perspective(900px) rotateX(${y}deg) rotateY(${x}deg)`;
      });
      item.addEventListener("pointerleave", () => { item.style.transform = ""; });
    });

    document.querySelectorAll(".magnetic").forEach((item) => {
      item.addEventListener("pointermove", (event) => {
        const rect = item.getBoundingClientRect();
        item.style.transform = `translate(${(event.clientX - rect.left - rect.width / 2) * .12}px, ${(event.clientY - rect.top - rect.height / 2) * .12}px)`;
      });
      item.addEventListener("pointerleave", () => { item.style.transform = ""; });
    });

    document.querySelectorAll("[data-parallax]").forEach((item) => {
      window.addEventListener("scroll", () => {
        item.style.transform = `translateY(${window.scrollY * Number(item.dataset.parallax)}px)`;
      }, { passive: true });
    });
  }

  const canvas = document.querySelector("#particle-canvas");
  if (!(canvas instanceof HTMLCanvasElement) || reducedMotion) return;
  const context = canvas.getContext("2d");
  const particles = [];
  function resizeCanvas() {
    canvas.width = window.innerWidth * devicePixelRatio;
    canvas.height = window.innerHeight * devicePixelRatio;
    canvas.style.width = `${window.innerWidth}px`;
    canvas.style.height = `${window.innerHeight}px`;
    context.setTransform(devicePixelRatio, 0, 0, devicePixelRatio, 0, 0);
    particles.length = 0;
    const count = Math.min(45, Math.floor(window.innerWidth / 28));
    for (let i = 0; i < count; i += 1) {
      particles.push({
        x: Math.random() * window.innerWidth,
        y: Math.random() * window.innerHeight,
        radius: Math.random() * 1.5 + .35,
        speed: Math.random() * .18 + .06,
        opacity: Math.random() * .45 + .12
      });
    }
  }
  function drawParticles() {
    context.clearRect(0, 0, window.innerWidth, window.innerHeight);
    particles.forEach((particle) => {
      particle.y -= particle.speed;
      if (particle.y < -5) particle.y = window.innerHeight + 5;
      context.beginPath();
      context.arc(particle.x, particle.y, particle.radius, 0, Math.PI * 2);
      context.fillStyle = `rgba(34, 211, 238, ${particle.opacity})`;
      context.fill();
    });
    requestAnimationFrame(drawParticles);
  }
  resizeCanvas();
  window.addEventListener("resize", resizeCanvas);
  requestAnimationFrame(drawParticles);
})();
