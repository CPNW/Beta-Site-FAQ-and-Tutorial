(function () {
  const quotes = [
    { text: "It is not how much you do, but how much love you put into doing.", author: "Mother Teresa" },
    { text: "Rejoice in your work; never lose sight of the nursing leader you are now and the nursing leader you will become.", author: "Sue Fitzsimons" },
    { text: "Never give up on a dream just because of the time it will take to accomplish it. The time will pass anyway.", author: "Earl Nightingale" },
    { text: "Work for a cause, not for applause. Live life to express, not to impress. Don't strive to make your presence noticed, just make your presence felt.", author: "Grace Lichtenstein" },
    { text: "You treat a disease: You win, you lose. You treat a person, I guarantee you win - no matter the outcome.", author: 'Hunter Doherty "Patch" Adams' },
    { text: "Every nurse was drawn to nursing because of a desire to care, to serve, or to help.", author: "Christina Feist-Heilmeier, RN" },
    { text: "Doctors save lives; nurses save doctors.", author: "Unknown" },
    { text: "Patients don't care about your policies. They care about how you treat them.", author: "Barbara Balik, healthcare consultant and author" },
    { text: "Nurses are the hospitality of the hospital.", author: "Carrie Latet" },
    { text: "You treat a disease: you win, you lose. You treat a person, I guarantee you win...no matter the outcome.", author: "Robin Williams, Patch Adams" },
    { text: "Just so you know, doctors are here to diagnose, not heal. We heal.", author: "Edie Falco in Nurse Jackie" },
    { text: "Bound by paperwork, short on hands, sleep, and energy...nurses are rarely short on caring.", author: "Sharon Hudacek" },
    { text: "Most of all, you need moral courage because nursing is about the pursuit of justice. It requires you to stand up to bullies, to do things that are right but difficult, and to speak your mind even when you are afraid. I wasn't strong like this when I started out. Nursing made me strong.", author: "Tilda Shalof" },
    { text: "In healthcare, the experience of the patient is the new marketing.", author: "David Feinberg, former CEO of Geisinger Health" },
    { text: "To do what nobody else will do, a way that nobody else can do, in spite of all we go through; that is to be a nurse.", author: "Rawsi Williams" },
    { text: "Happiness ... consists of giving, and in serving others.", author: "Henry Drummond" },
    { text: "Have a heart that never hardens, a temper that never tires, a touch that never hurts.", author: "Charles Dickens" },
    { text: "A nurse is compassion in scrubs.", author: "Lexie Saige" },
    { text: "To truly improve the patient experience, we must understand the patient journey from the patient's perspective.", author: "Jason A. Wolf, President of The Beryl Institute" },
    { text: "I attribute my success to this: I never gave or took any excuse.", author: "Florence Nightingale" }
  ];

  document.querySelectorAll('[data-current-year]').forEach((el) => {
    el.textContent = new Date().getFullYear();
  });

  const quoteTarget = document.querySelector('[data-inspiration-quote]');
  if (quoteTarget && quotes.length) {
    const pick = quotes[Math.floor(Math.random() * quotes.length)];
    quoteTarget.innerHTML = `“${pick.text}” <span class="d-block mt-2">&mdash; ${pick.author}</span>`;
  }
})();
