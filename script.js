const weddingDate = new Date("2026-10-09T00:00:00-05:00");

const countdownElements = {
  days: document.getElementById("countdown-days"),
  hours: document.getElementById("countdown-hours"),
  minutes: document.getElementById("countdown-minutes"),
  seconds: document.getElementById("countdown-seconds")
};

function updateCountdown() {
  const remaining = Math.max(0, weddingDate.getTime() - Date.now());
  const totalSeconds = Math.floor(remaining / 1000);

  const days = Math.floor(totalSeconds / 86400);
  const hours = Math.floor((totalSeconds % 86400) / 3600);
  const minutes = Math.floor((totalSeconds % 3600) / 60);
  const seconds = totalSeconds % 60;

  countdownElements.days.textContent = String(days).padStart(2, "0");
  countdownElements.hours.textContent = String(hours).padStart(2, "0");
  countdownElements.minutes.textContent = String(minutes).padStart(2, "0");
  countdownElements.seconds.textContent = String(seconds).padStart(2, "0");
}

updateCountdown();
setInterval(updateCountdown, 1000);

const backgroundMusic = document.getElementById("background-music");
const musicToggle = document.getElementById("music-toggle");
let musicHasStarted = false;

function removeMusicStartListeners() {
  document.removeEventListener("pointerdown", startBackgroundMusic);
  document.removeEventListener("touchstart", startBackgroundMusic);
  document.removeEventListener("keydown", startBackgroundMusic);
}

function updateMusicToggle() {
  const isPlaying = !backgroundMusic.paused || !musicHasStarted;
  musicToggle.classList.toggle("music-toggle-paused", !isPlaying);
  musicToggle.setAttribute("aria-pressed", String(isPlaying));
  musicToggle.setAttribute(
    "aria-label",
    isPlaying ? "Pausar música" : "Reproducir música"
  );
}

function startBackgroundMusic(event) {
  if (event?.target?.closest?.("#music-toggle")) {
    return;
  }

  const playback = backgroundMusic.play();

  if (playback) {
    playback
      .then(() => {
        musicHasStarted = true;
        removeMusicStartListeners();
        updateMusicToggle();
      })
      .catch(() => {
        // iOS y Android pueden exigir el primer toque del visitante.
        // El control permanece activo y ese toque inicia la música.
      });
  }
}

document.addEventListener("pointerdown", startBackgroundMusic, {
  passive: true
});
document.addEventListener("touchstart", startBackgroundMusic, {
  passive: true
});
document.addEventListener("keydown", startBackgroundMusic);

startBackgroundMusic();

musicToggle.addEventListener("click", () => {
  if (backgroundMusic.paused) {
    backgroundMusic
      .play()
      .then(() => {
        musicHasStarted = true;
        removeMusicStartListeners();
        updateMusicToggle();
      })
      .catch(updateMusicToggle);
  } else {
    backgroundMusic.pause();
  }
});

backgroundMusic.addEventListener("play", () => {
  musicHasStarted = true;
  updateMusicToggle();
});
backgroundMusic.addEventListener("pause", updateMusicToggle);
updateMusicToggle();

const openingVideo = document.getElementById("opening-video");
let videoRetryCount = 0;
let videoRetryTimer;

function playOpeningVideo() {
  if (!openingVideo) {
    return;
  }

  openingVideo.muted = true;
  openingVideo.defaultMuted = true;
  openingVideo.playsInline = true;
  openingVideo.play().catch(() => {
    // El póster permanece visible mientras el navegador termina de cargar.
  });
}

function retryOpeningVideo() {
  if (!openingVideo || videoRetryCount >= 3) {
    return;
  }

  window.clearTimeout(videoRetryTimer);
  videoRetryCount += 1;
  videoRetryTimer = window.setTimeout(() => {
    openingVideo.load();
    playOpeningVideo();
  }, videoRetryCount * 1200);
}

openingVideo?.addEventListener("loadeddata", () => {
  videoRetryCount = 0;
  playOpeningVideo();
});
openingVideo?.addEventListener("canplay", playOpeningVideo);
openingVideo?.addEventListener("stalled", retryOpeningVideo);
openingVideo?.addEventListener("error", retryOpeningVideo);
openingVideo?.addEventListener("ended", playOpeningVideo);

document.addEventListener("visibilitychange", () => {
  if (!document.hidden) {
    playOpeningVideo();
  }
});

window.addEventListener("pageshow", playOpeningVideo);
playOpeningVideo();

const carouselImage = document.getElementById("carousel-image");
const carouselFrame = document.querySelector(".carousel-frame");
const carouselPrevious = document.getElementById("carousel-previous");
const carouselNext = document.getElementById("carousel-next");
const carouselDots = Array.from(document.querySelectorAll(".carousel-dot"));
const carouselPhotos = [
  "img/carrusel/1.jpg",
  "img/carrusel/2.jpg",
  "img/carrusel/3.jpg",
  "img/carrusel/4.jpg",
  "img/carrusel/5.jpg",
  "img/carrusel/6.jpg",
  "img/carrusel/7.jpg"
];
let currentCarouselSlide = 0;
let carouselTouchStartX = 0;
let carouselTouchStartY = 0;

function showCarouselSlide(index) {
  currentCarouselSlide =
    (index + carouselPhotos.length) % carouselPhotos.length;

  carouselImage.src = carouselPhotos[currentCarouselSlide];
  carouselImage.alt =
    `Fotografía ${currentCarouselSlide + 1} de Jadhira y Carlos`;
  carouselImage.classList.toggle(
    "carousel-image-monochrome",
    currentCarouselSlide % 2 === 0
  );

  carouselDots.forEach((dot, dotIndex) => {
    const isActive = dotIndex === currentCarouselSlide;
    dot.classList.toggle("carousel-dot-active", isActive);
    if (isActive) {
      dot.setAttribute("aria-current", "true");
    } else {
      dot.removeAttribute("aria-current");
    }
  });
}

carouselPrevious.addEventListener("click", () => {
  showCarouselSlide(currentCarouselSlide - 1);
});

carouselNext.addEventListener("click", () => {
  showCarouselSlide(currentCarouselSlide + 1);
});

carouselDots.forEach((dot) => {
  dot.addEventListener("click", () => {
    showCarouselSlide(Number(dot.dataset.slide));
  });
});

carouselFrame.addEventListener(
  "touchstart",
  (event) => {
    const touch = event.touches[0];
    carouselTouchStartX = touch.clientX;
    carouselTouchStartY = touch.clientY;
  },
  { passive: true }
);

carouselFrame.addEventListener(
  "touchend",
  (event) => {
    const touch = event.changedTouches[0];
    const horizontalDistance = touch.clientX - carouselTouchStartX;
    const verticalDistance = touch.clientY - carouselTouchStartY;
    const isHorizontalSwipe =
      Math.abs(horizontalDistance) >= 45 &&
      Math.abs(horizontalDistance) > Math.abs(verticalDistance) * 1.2;

    if (!isHorizontalSwipe) {
      return;
    }

    if (horizontalDistance < 0) {
      showCarouselSlide(currentCarouselSlide + 1);
    } else {
      showCarouselSlide(currentCarouselSlide - 1);
    }
  },
  { passive: true }
);

const revealTitles = document.querySelectorAll(
  [
    "#wedding-message-title",
    ".parents-group h2",
    ".save-date h2",
    ".marriage-counsel h3",
    ".reception h3",
    ".itinerary-section h2",
    ".dress-code-section h2",
    ".rsvp-section h2",
    ".invitation-closing"
  ].join(",")
);

revealTitles.forEach((title) => {
  title.classList.add("scroll-reveal");
});

if (
  "IntersectionObserver" in window &&
  !window.matchMedia("(prefers-reduced-motion: reduce)").matches
) {
  const revealObserver = new IntersectionObserver(
    (entries) => {
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          entry.target.classList.add("scroll-reveal-visible");
        }
      });
    },
    {
      threshold: 0.15,
      rootMargin: "0px 0px -4% 0px"
    }
  );

  revealTitles.forEach((title) => {
    revealObserver.observe(title);
  });
} else {
  revealTitles.forEach((title) => {
    title.classList.add("scroll-reveal-visible");
  });
}

let hasLeftInvitationTop = false;
const closingMessage = document.querySelector(".invitation-closing");

function manageRevealCycle() {
  const pageTop = window.scrollY <= 8;

  if (!pageTop) {
    hasLeftInvitationTop = true;
  } else if (hasLeftInvitationTop) {
    revealTitles.forEach((title) => {
      title.classList.remove("scroll-reveal-visible");
    });
    hasLeftInvitationTop = false;
  }

  const reachedPageEnd =
    window.scrollY + window.innerHeight >=
    document.documentElement.scrollHeight - 8;

  if (reachedPageEnd) {
    closingMessage?.classList.add("scroll-reveal-visible");
  }
}

window.addEventListener("scroll", manageRevealCycle, { passive: true });
manageRevealCycle();
