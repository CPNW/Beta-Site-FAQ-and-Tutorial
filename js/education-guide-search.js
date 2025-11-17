document.addEventListener("DOMContentLoaded", () => {
  const searchInput = document.getElementById("educationGuideSearchInput");
  const resultsContainer = document.getElementById("educationGuideSearchResults");
  const guideColumns = Array.from(document.querySelectorAll("[data-guide-id]"));
  const MIN_TERM_LENGTH = 2;

  if (!searchInput || !resultsContainer || !guideColumns.length) return;

  const guides = guideColumns.map((column) => {
    const id = column.dataset.guideId;
    const title = column.querySelector(".card h2, .card .h5")?.textContent.trim() ?? "";
    const description = column.querySelector(".card p")?.textContent.trim() ?? "";
    const url = column.querySelector("a")?.getAttribute("href") ?? "#";
    const tags = (column.dataset.guideTags ?? "")
      .split(",")
      .map((tag) => tag.trim())
      .filter(Boolean);

    return { id, column, title, description, url, tags }; 
  });

  const showMessage = (message) => {
    resultsContainer.innerHTML = `<p class="text-muted small mb-0">${message}</p>`;
  };

  const toggleResultVisibility = (show) => {
    resultsContainer.classList.toggle("faq-match-hidden", !show);
  };

  const renderResults = (matches) => {
    if (!matches.length) {
      showMessage("No guides match that term. Try another keyword.");
      toggleResultVisibility(true);
      return;
    }

    resultsContainer.innerHTML = matches
      .map(
        (guide) => `
          <a class="faq-result-card" href="${guide.url}">
            <h3>${guide.title}</h3>
            <p class="small mb-0">${guide.description}</p>
            <div class="faq-result-tags">
              ${guide.tags.map((tag) => `<span class="faq-result-tag">${tag}</span>`).join("")}
            </div>
          </a>
        `
      )
      .join("");
    toggleResultVisibility(true);
  };

  const handleSearch = (value) => {
    const term = value.trim().toLowerCase();

    if (term.length < MIN_TERM_LENGTH) {
      toggleResultVisibility(false);
      showMessage("Start typing to filter guides or jump to a specific article.");
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
