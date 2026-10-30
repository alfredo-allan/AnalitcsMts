// js/home.js - Animações GSAP e Controle de Tema

document.addEventListener("DOMContentLoaded", () => {
  // 1. REGISTRO DO GSAP SCROLLTRIGGER
  gsap.registerPlugin(ScrollTrigger);

  // 2. ANIMAÇÃO DE ENTRADA DO HERO (FADE + SLIDE UP)
  gsap.from(".hero-anim", {
    opacity: 0,
    y: 40,
    duration: 0.9,
    stagger: 0.15,
    ease: "power3.out",
  });

  // 3. EFEITO PARALLAX SUAVE NO MOCKUP CARD CONFORME SCROLL
  gsap.to(".hero-mockup-card", {
    scrollTrigger: {
      trigger: ".hero-section",
      start: "top top",
      end: "bottom top",
      scrub: 1,
    },
    y: 60,
    scale: 0.98,
    opacity: 0.85,
  });

  // 4. ANIMAÇÃO DOS FEATURE CARDS COM SCROLL TRIGGER
  gsap.from(".feature-card-item", {
    scrollTrigger: {
      trigger: "#features",
      start: "top 80%",
      toggleActions: "play none none reverse",
    },
    opacity: 0,
    y: 50,
    duration: 0.8,
    stagger: 0.2,
    ease: "power2.out",
  });

  // 5. LÓGICA DE TOGGLE DE TEMA (CLARO / ESCURO)
  const themeToggleBtn = document.getElementById("themeToggleBtn");
  const themeIcon = document.getElementById("themeIcon");
  const themeLabel = document.getElementById("themeLabel");

  if (themeToggleBtn) {
    themeToggleBtn.addEventListener("click", () => {
      const currentTheme = document.documentElement.getAttribute("data-theme");
      const newTheme = currentTheme === "dark" ? "light" : "dark";

      document.documentElement.setAttribute("data-theme", newTheme);

      if (newTheme === "dark") {
        if (themeLabel) themeLabel.innerText = "Modo Claro";
        if (themeIcon) themeIcon.className = "fa-solid fa-sun";
      } else {
        if (themeLabel) themeLabel.innerText = "Modo Escuro";
        if (themeIcon) themeIcon.className = "fa-solid fa-moon";
      }
    });
  }
});
