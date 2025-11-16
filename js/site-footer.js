(function () {
  const quotes = [
    { text: "Wherever the art of Medicine is loved, there is also a love of Humanity.", author: "Hippocrates" },
    { text: "I attribute my success to this: I never gave or took any excuse.", author: "Florence Nightingale" },
    { text: "The good physician treats the disease; the great physician treats the patient who has the disease.", author: "William Osler" },
    { text: "Chance favors only the prepared mind.", author: "Louis Pasteur" },
    { text: "I may be compelled to face danger, but never fear it, and while our soldiers can stand and fight, I can stand and feed and nurse them.", author: "Clara Barton" },
    { text: "If society will not admit of woman's free development, then society must be remodeled.", author: "Elizabeth Blackwell" },
    { text: "Medicine is a social science, and politics is nothing else but medicine on a large scale.", author: "Rudolf Virchow" },
    { text: "Childbed fever is caused by decomposed organic matter conveyed to the womb.", author: "Ignaz Semmelweis" },
    { text: "An operation ought not to be performed for the surgeon's satisfaction but for the patient's benefit.", author: "Joseph Lister" },
    { text: "Let the generations know that women in uniform also guaranteed their freedom.", author: "Mary Edwards Walker" }
  ];

  document.querySelectorAll('[data-current-year]').forEach((el) => {
    el.textContent = new Date().getFullYear();
  });

  const quoteTarget = document.querySelector('[data-inspiration-quote]');
  if (quoteTarget && quotes.length) {
    const pick = quotes[Math.floor(Math.random() * quotes.length)];
    quoteTarget.innerHTML = `“${pick.text}” <span class="d-block mt-2">— ${pick.author}</span>`;
  }
})();
