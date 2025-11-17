document.addEventListener("DOMContentLoaded", () => {
  const searchInput = document.getElementById("studentGuideSearchInput");
  const resultsContainer = document.getElementById("studentGuideSearchResults");
  const guideCards = Array.from(document.querySelectorAll("[data-student-guide]"));
  const MIN_TERM_LENGTH = 2;

  if (!searchInput || !resultsContainer || !guideCards.length) return;

  const guides = guideCards.map((card) => {
    const id = card.dataset.studentGuide;
    const title = card.querySelector(".card h2, .card .h5")?.textContent.trim() ?? "";
    const description = card.querySelector(".card p")?.textContent.trim() ?? "";
    const url = card.querySelector("a")?.getAttribute("href") ?? "#";
    const tags = (card.dataset.guideTags ?? "")
      .split(",")
      .map((tag) => tag.trim())
      .filter(Boolean);

    return { id, title, description, url, tags };
  });

  const toggleResults = (show) => {
    resultsContainer.classList.toggle("faq-match-hidden", !show);
  };

  const showMessage = (message) => {
    resultsContainer.innerHTML = `<p class="text-muted small mb-0">${message}</p>`;
  };

  const renderResults = (matches) => {
    if (!matches.length) {
      showMessage("No guides match that term. Try another keyword.");
      toggleResults(true);
      return;
    }

    resultsContainer.innerHTML = matches
      .map(
        (guide) => `
          <a class="faq-result-card" href="${guide.url}">
            <h3>${guide.title}</h3>
            <p class="small mb-0">${guide.description}</p>
            <div class="faq-result-tags">
              ${guide.tags.map((tag) => `<span class=\"faq-result-tag\">${tag}</span>`).join("")}
            </div>
          </a>
        `
      )
      .join("");
    toggleResults(true);
  };

  const handleSearch = (value) => {
    const term = value.trim().toLowerCase();

    if (term.length < MIN_TERM_LENGTH) {
      toggleResults(false);
      showMessage("Start typing to filter student guides.");
      return;
    }

    const matches = guides.filter((guide) => {
      return (
        guide.title.toLowerCase().includes(term) ||
        guide.description.toLowerCase().includes(term) ||
        guide.tags.some((tag) => tag.toLowerCase().includes(term))
      );
    });

    renderResults(matches);
  };

  searchInput.addEventListener("input", (event) => handleSearch(event.target.value));
  handleSearch("");
});
