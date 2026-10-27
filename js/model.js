/* =====================================================================
   MODEL — the data and state of the app. Never touches the DOM.
   Portofolio Bisnis SANZ GOKU — Muhammad Sandika Akbar
   ===================================================================== */

const Model = {

  githubUser: "Sanzxgoku",

  // Where the contact form delivers
  contactEmail: "admin@snz.my.id",

  // App state (read/written by the Controller, displayed by the View)
  state: {
    screen: "home",        // which screen is showing
    menuIndex: 0,          // selected item on the home menu
    reposLoaded: false,
    skillsBuilt: false,
  },

  // ---- Featured services & products ----
  featured: [
    {
      title: "VPS & Layanan Server Linux",
      tag: "Server & Cloud", color: "#e60012", live: true,
      url: "https://snz.my.id", cta: "Kunjungi snz.my.id →",
      img: "assets/projects/vps.png",
      desc: "Layanan penyediaan Virtual Private Server (VPS) berbasis Linux dengan performa andal, konfigurasi server yang optimal, serta pemeliharaan infrastruktur untuk kebutuhan hosting dan aplikasi.",
    },
    {
      title: "Panel Pterodactyl & Hosting",
      tag: "Game/App Panel", color: "#3178c6", live: true,
      url: "https://t.me/Sanzxgoku", cta: "Pesan via Telegram →",
      img: "assets/projects/ptero.png",
      desc: "Instalasi, konfigurasi, dan penyediaan panel Pterodactyl untuk pengelolaan server game, bot, dan aplikasi mandiri dengan kontrol resource yang fleksibel dan mudah digunakan.",
    },
    {
      title: "Layanan Nokos & Verifikasi OTP",
      tag: "Digital Service", color: "#f1e05a", live: true,
      url: "https://t.me/Sanzstrnokosmumerbot", cta: "Order via Bot OTP →",
      img: "assets/projects/otp.png",
      desc: "Penyediaan nomor virtual (nokos) untuk kebutuhan verifikasi OTP berbagai platform digital secara instan via Bot @Sanzstrnokosmumerbot, didukung integrasi RumahOTP.",
    },
    {
      title: "Bot Telegram: @Sanzstrnokosmumerbot",
      tag: "Automation", color: "#3dff6e", live: true,
      url: "https://t.me/Sanzstrnokosmumerbot", cta: "Buka Bot Telegram →",
      img: "assets/projects/bot.png",
      desc: "Sistem bot Telegram resmi SANZ GOKU (@Sanzstrnokosmumerbot) dengan moto 'Satu Bot, Banyak Solusi' untuk otomatisasi katalog produk, layanan nomor OTP, cek status pesanan, dan transaksi.",
    },
  ],

  // Repos already shown in "featured" get hidden from the GitHub feed
  featuredRepoNames: [
    "vps-management",
    "pterodactyl-setup",
    "telegram-bot-core",
    "nokos-service",
  ],

  // Shown if the GitHub API can't be reached or for offline demo
  fallbackRepos: [
    {
      name: "SOTP (sxgotp) - OTP Web Platform", language: "Next.js / React", stargazers_count: 0,
      html_url: "https://github.com/Omicron69",
      description: "Proyek pengembangan website layanan OTP mandiri menggunakan Next.js, Tailwind CSS, Vercel, dan Supabase. Berstatus sebagai proyek pengembangan tersendiri.",
    },
    {
      name: "Telegram-Business-Bot", language: "JavaScript", stargazers_count: 0,
      html_url: "https://github.com/Omicron69",
      description: "Sistem bot Telegram berbasis Node.js dan Telegram Bot API untuk menampilkan katalog layanan, menangani pesanan pelanggan, dan integrasi API.",
    },
    {
      name: "Digital-Products-Subdomain-Manager", language: "Node.js", stargazers_count: 0,
      html_url: "https://github.com/Omicron69",
      description: "Sistem pendukung pengelolaan produk digital, layanan domain/subdomain, serta integrasi Cloudflare API untuk kebutuhan pelanggan.",
    },
    {
      name: "XS - Personal AI Assistant (Research)", language: "TypeScript", stargazers_count: 0,
      html_url: "https://github.com/Omicron69",
      description: "Konsep riset asisten AI pribadi (analisis file, riset web, smart note). Proyek teknologi terpisah di luar lini bisnis utama SANZ GOKU.",
    },
  ],

  // Optional thumbnail overrides
  projectImages: {},

  langColors: {
    JavaScript: "#f1e05a", TypeScript: "#3178c6", Python: "#3572A5",
    "Next.js / React": "#000000", "Node.js": "#68a063", Linux: "#f34b7d",
    HTML: "#e34c26", CSS: "#663399",
  },

  // ---- Skills screen: Core Competencies & Technologies ----
  skills: [
    { group: "Server & Infrastruktur", items: [
      ["Linux VPS (Ubuntu / Debian)", 90], ["Panel Pterodactyl Management", 88],
      ["Cloudflare DNS & CDN", 85], ["DigitalOcean / Cloud Infrastructure", 80],
      ["Server Security & Maintenance", 84], ["Port & Firewall Configuration", 82],
    ]},
    { group: "Bot & Pengembangan Sistem", items: [
      ["JavaScript / Node.js", 86], ["Telegram Bot API", 92],
      ["node-telegram-bot-api", 88], ["REST API & Axios Integration", 85],
      ["Next.js / React / Supabase (SOTP)", 78], ["Automation & Scripting", 84],
    ]},
    { group: "Layanan & Integrasi Pihak Ketiga", items: [
      ["RumahOTP API Integration", 88], ["Pakasir QRIS Payment Gateway", 82],
      ["FayuPedia API Integration", 80], ["Subdomain & DNS Management", 85],
      ["Boost Medsos & Digital Products", 84], ["Top Up Saldo Digital", 86],
    ]},
    { group: "Operasional & Pelayanan Bisnis", items: [
      ["Customer Service & Fast Response", 94], ["Alur Transaksi & Verifikasi Order", 90],
      ["Komunikasi Transparan & Jujur", 95], ["Evaluasi & Riset Kebutuhan Pasar", 85],
    ]},
  ],

  // ---- Data fetching ----
  async fetchRepos() {
    return { repos: this.fallbackRepos, live: false };
  },
};
