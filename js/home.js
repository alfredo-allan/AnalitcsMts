// js/home.js - Animações GSAP Refinadas & Controle de Tema

document.addEventListener("DOMContentLoaded", () => {
  // 1. REGISTRO DO GSAP SCROLLTRIGGER
  gsap.registerPlugin(ScrollTrigger);

  // 2. ANIMAÇÃO DE ENTRADA DO HERO (FADE + SLIDE UP SUAVE AO CARREGAR)
  gsap.from(".hero-anim", {
    opacity: 0,
    y: 35,
    duration: 0.9,
    stagger: 0.12,
    ease: "power3.out",
  });

  // 3. PARALLAX SUAVE NO MOCKUP CARD (SEM SUBIR POR CIMA DOS BOTÕES)
  // Usamos scrub com suavização (smoothness) e ajustamos a faixa de scroll
  gsap.to(".hero-mockup-card", {
    scrollTrigger: {
      trigger: ".hero-section",
      start: "top 10%", // Começa um pouco depois do topo para não engasgar
      end: "bottom top", // Vai até o fim da seção hero
      scrub: 1.2, // Suaviza o movimento acompanhando o scroll naturalmente
      invalidateOnRefresh: true,
    },
    y: 40, // Move para baixo suavemente sem invadir a área dos botões
    scale: 0.97, // Efeito sutil de profundidade
    opacity: 0.9,
  });

  // 4. ANIMAÇÃO REFINADA DOS FEATURE CARDS COM REVELAÇÃO NATIVA
  gsap.from(".feature-card-item", {
    scrollTrigger: {
      trigger: "#features",
      start: "top 75%", // Dispara quando 25% da seção de recursos entra na tela
      toggleActions: "play none none reverse",
    },
    opacity: 0,
    y: 45,
    duration: 0.7,
    stagger: 0.18,
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
