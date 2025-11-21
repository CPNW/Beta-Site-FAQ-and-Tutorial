document.addEventListener("DOMContentLoaded", () => {
  const searchInput = document.getElementById("faqSearchInput");
  const resultsContainer = document.getElementById("faqSearchResults");
  const accordion = document.getElementById("studentFaqAccordion");
  const MIN_TERM_LENGTH = 2;

  if (!searchInput || !resultsContainer || !accordion) return;

  const faqEntries = Array.from(accordion.querySelectorAll(".accordion-item")).map((item) => {
    const question = item.querySelector(".accordion-button")?.textContent ?? "";
    const answer = item.querySelector(".accordion-body")?.textContent ?? "";
    return {
      element: item,
      content: `${question} ${answer}`.toLowerCase(),
    };
  });

  const articleLinks = [
    {
      title: "CPNW Requirements Checklist",
      description: "Step-by-step guide for WATCH background checks, immunization uploads, and other compliance evidence.",
      url: "student-cpnw-requirements.html",
      tags: ["requirements", "watch", "background check", "documentation"],
    },
    {
      title: "Student Dashboard Walkthrough",
      description: "Learn how to interpret notifications, track placement tasks, and stay ahead on deadlines.",
      url: "student-dashboard.html",
      tags: ["dashboard", "notifications", "alerts", "status"],
    },
    {
      title: "eLearning Modules Overview",
      description: "See the 10 mandatory modules, completion tips, and how to confirm each module is showing as completed.",
      url: "student-elearning-modules.html",
      tags: ["elearning", "modules", "training", "courses"],
    },
    {
      title: "How to Register with CPNW",
      description: "New to CPNW? Follow this tutorial to create your account, verify demographics, and set up security settings.",
      url: "how-to-register.html",
      tags: ["register", "account", "login", "setup"],
    },
    {
      title: "Understanding Requirement Statuses",
      description: "Explains the color coding, pending review state, and what to do if something is red or expired.",
      url: "student-requirements.html",
      tags: ["status", "pending", "expired", "upload"],
    },
    {
      title: "Accepted Documentation Formats",
      description: "See required data points, approved file types, and examples of acceptable uploads before you submit.",
      url: "accepted-document-formats.html",
      tags: ["documents", "uploads", "formats", "examples"],
    },
  ];

  const renderArticleResults = (term) => {
    const trimmedTerm = term.trim();

    if (trimmedTerm.length < MIN_TERM_LENGTH) {
      resultsContainer.innerHTML =
        '<p class="text-muted small mb-0">Start typing to see suggested guides and tutorials.</p>';
      return;
    }

    const matches = articleLinks.filter((article) => {
      const needle = trimmedTerm.toLowerCase();
      return (
        article.title.toLowerCase().includes(needle) ||
        article.description.toLowerCase().includes(needle) ||
        article.tags.some((tag) => tag.toLowerCase().includes(needle))
      );
    });

    if (!matches.length) {
      resultsContainer.innerHTML =
        '<p class="text-muted small mb-0">No related guides found. Try different keywords.</p>';
      return;
    }

    resultsContainer.innerHTML = matches
      .map(
        (article) => `
        <a class="faq-result-card" href="${article.url}">
          <h3>${article.title}</h3>
          <p class="small">${article.description}</p>
          <div class="faq-result-tags">
            ${article.tags
              .map((tag) => `<span class="faq-result-tag">${tag}</span>`)
              .join("")}
          </div>
        </a>
      `
      )
      .join("");
  };

  const toggleFaqVisibility = (term) => {
    const normalizedTerm = term.trim().toLowerCase();
    const showAll = normalizedTerm.length < MIN_TERM_LENGTH;

    faqEntries.forEach((entry) => {
      const shouldShow = showAll || entry.content.includes(normalizedTerm);
      entry.element.classList.toggle("faq-match-hidden", !shouldShow);
    });
  };

  searchInput.addEventListener("input", (event) => {
    const term = event.target.value;
    toggleFaqVisibility(term);
    renderArticleResults(term);
  });

  renderArticleResults("");
});
