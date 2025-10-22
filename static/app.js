document.addEventListener("DOMContentLoaded", function() {
  // Загрузка дополнительных квестов
  const loadBtn = document.getElementById("load-more");
  if (loadBtn) {
    loadBtn.addEventListener("click", async () => {
      const btn = loadBtn;
      const skip = parseInt(btn.dataset.skip || "0");
      const params = new URLSearchParams(window.location.search);
      params.set("skip", skip);

      const url = "/api/quests?" + params.toString();
      btn.disabled = true;
      btn.textContent = "Загрузка...";

      try {
        const res = await fetch(url);
        if (res.ok) {
          const html = await res.text();
          const tempDiv = document.createElement('div');
          tempDiv.innerHTML = html;
          const newGrid = tempDiv.querySelector('.cards-grid');

          const cardsGrid = document.querySelector(".cards-grid");
          if (newGrid && cardsGrid) {
            // Добавляем новые карточки
            const newCards = newGrid.innerHTML;
            cardsGrid.insertAdjacentHTML('beforeend', newCards);

            const addedCount = (newGrid.innerHTML.match(/class="card"/g) || []).length;
            const newSkip = skip + addedCount;
            btn.dataset.skip = newSkip;
            btn.textContent = "Посмотреть ещё";
            btn.disabled = false;

            // Скрываем кнопку если загружено меньше чем лимит
            if (addedCount === 0 || addedCount < 15) {
              btn.style.display = 'none';
            }
          } else {
            btn.textContent = "Нет данных";
            btn.disabled = true;
          }
        } else {
          btn.textContent = "Ошибка загрузки";
          btn.disabled = false;
        }
      } catch (error) {
        btn.textContent = "Ошибка";
        btn.disabled = false;
        console.error('Load more error:', error);
      }
    });
  }

  // Фильтры
  const applyFilters = document.getElementById("apply-filters");
  if (applyFilters) {
    applyFilters.addEventListener("click", () => {
      const form = document.getElementById("filter-form");
      const data = new FormData(form);
      const params = new URLSearchParams();
      for (const [k,v] of data.entries()) {
        if (v) params.set(k, v);
      }
      window.location = "/?" + params.toString();
    });
  }

  // Звезды рейтинга страха
  const fearStars = document.querySelectorAll("#fear-level span");
  const fearInput = document.getElementById("fear_input");

  fearStars.forEach(star => {
    star.addEventListener("click", () => {
      const value = parseInt(star.getAttribute("data-value"));
      fearInput.value = value;

      // Обновляем визуальное отображение
      fearStars.forEach((s, index) => {
        if (index < value) {
          s.classList.add("active");
        } else {
          s.classList.remove("active");
        }
      });
    });
  });

  // Выбор количества игроков
  const playerCircles = document.querySelectorAll("#players span");
  const playersInput = document.getElementById("players_input");

  playerCircles.forEach(circle => {
    circle.addEventListener("click", () => {
      const value = parseInt(circle.getAttribute("data-value"));
      playersInput.value = value;

      // Обновляем визуальное отображение
      playerCircles.forEach((c, index) => {
        if (index < value) {
          c.classList.add("active");
        } else {
          c.classList.remove("active");
        }
      });
    });
  });

  // Сортировка
  const sortSelect = document.getElementById('sort-select');
  if (sortSelect) {
    sortSelect.addEventListener('change', function() {
      const url = new URL(window.location);
      url.searchParams.set('sort', this.value);
      window.location.href = url.toString();
    });
  }

  // Инициализация фильтров из URL параметров
  function initFiltersFromURL() {
    const urlParams = new URLSearchParams(window.location.search);

    // Уровень страха
    const fearLevel = urlParams.get('fear_level');
    if (fearLevel && fearInput) {
      fearInput.value = fearLevel;
      fearStars.forEach((star, index) => {
        if (index < fearLevel) {
          star.classList.add("active");
        }
      });
    }

    // Количество игроков
    const players = urlParams.get('players');
    if (players && playersInput) {
      playersInput.value = players;
      playerCircles.forEach((circle, index) => {
        if (index < players) {
          circle.classList.add("active");
        }
      });
    }

    // Чекбоксы жанров и сложности
    const genres = urlParams.getAll('genre');
    const difficulties = urlParams.getAll('difficulty');

    genres.forEach(genre => {
      const checkbox = document.querySelector(`input[name="genre"][value="${genre}"]`);
      if (checkbox) checkbox.checked = true;
    });

    difficulties.forEach(difficulty => {
      const checkbox = document.querySelector(`input[name="difficulty"][value="${difficulty}"]`);
      if (checkbox) checkbox.checked = true;
    });
  }

  initFiltersFromURL();
});