document.addEventListener("DOMContentLoaded", () => {
  const searchInput = document.getElementById("educationFaqSearchInput");
  const resultsContainer = document.getElementById("educationFaqSearchResults");
  const accordion = document.getElementById("educationFaqAccordion");
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
      title: "Coordinator Dashboard Overview",
      description: "Tour notifications, quick links, and admin shortcuts on the landing page.",
      url: "education-articles/educational-dashboard.html",
      tags: ["dashboard", "notifications", "admin", "shortcuts"],
    },
    {
      title: "Approving Account Requests",
      description: "Step-by-step for reviewing pending student and faculty access in the Users tab.",
      url: "education-articles/approving-account-requests.html",
      tags: ["accounts", "users", "access", "pending"],
    },
    {
      title: "Assign Students to Sites",
      description: "How to use the Assignments table for individuals or groups before rotations go live.",
      url: "education-articles/assigning-students.html",
      tags: ["assign", "students", "sites", "assignments"],
    },
    {
      title: "Background + WATCH Reports",
      description: "Initiate Checkr and CPNW WATCH services and understand reviewer outcomes.",
      url: "education-articles/background-watch-reports.html",
      tags: ["background", "watch", "checkr", "reports"],
    },
    {
      title: "Create Education Requirements",
      description: "Build coordinator-owned requirements, attach instructions, and publish to programs.",
      url: "education-articles/create-education-requirement.html",
      tags: ["requirements", "forms", "site orientations", "insurance"],
    },
    {
      title: "Create an Insurance Requirement",
      description: "Step-by-step instructions for building liability insurance requirements that align with affiliation agreements.",
      url: "education-articles/create-insurance-requirement.html",
      tags: ["insurance", "liability", "requirements", "coverage"],
    },
    {
      title: "Create a License Requirement",
      description: "Collect faculty or graduate student license details and documents through the Licenses requirement category.",
      url: "education-articles/create-license-requirement.html",
      tags: ["licenses", "credentials", "faculty", "students"],
    },
    {
      title: "Review Submissions",
      description: "Use the Review table to inspect uploads, mark Verification Complete, and signal healthcare partners via the Verified column.",
      url: "education-articles/review-submissions.html",
      tags: ["review", "verification complete", "verified", "needs review"],
    },
  ];

  const renderArticleResults = (term) => {
    const trimmedTerm = term.trim();

    if (trimmedTerm.length < MIN_TERM_LENGTH) {
      resultsContainer.innerHTML =
        '<p class="text-muted small mb-0">Start typing to see suggested coordinator guides.</p>';
      return;
    }

    const needle = trimmedTerm.toLowerCase();
    const matches = articleLinks.filter(
      (article) =>
        article.title.toLowerCase().includes(needle) ||
        article.description.toLowerCase().includes(needle) ||
        article.tags.some((tag) => tag.toLowerCase().includes(needle))
    );

    if (!matches.length) {
      resultsContainer.innerHTML = '<p class="text-muted small mb-0">No related guides found. Try different keywords.</p>';
      return;
    }

    resultsContainer.innerHTML = matches
      .map(
        (article) => `
        <a class="faq-result-card" href="${article.url}">
          <h3>${article.title}</h3>
          <p class="small">${article.description}</p>
          <div class="faq-result-tags">
            ${article.tags.map((tag) => `<span class="faq-result-tag">${tag}</span>`).join("")}
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
