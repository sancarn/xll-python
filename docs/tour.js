const tour = document.querySelector(".tour");
const navigation = tour.querySelector(".tour-nav");
const tabs = [...navigation.querySelectorAll("a")];
const panels = tabs.map(tab => document.querySelector(tab.hash));

function selectTab(selected) {
  tabs.forEach((tab, index) => {
    const current = tab === selected;
    tab.setAttribute("aria-selected", String(current));
    tab.tabIndex = current ? 0 : -1;
    panels[index].hidden = !current;
  });
}

function selectFromHash() {
  const selected = tabs.find(tab => tab.hash === location.hash);
  if (selected) selectTab(selected);
}

navigation.setAttribute("role", "tablist");
tabs.forEach((tab, index) => {
  const panel = panels[index];
  tab.setAttribute("role", "tab");
  tab.setAttribute("aria-controls", panel.id);
  panel.setAttribute("role", "tabpanel");
  panel.setAttribute("aria-labelledby", tab.id);
  tab.addEventListener("click", event => {
    if (event.ctrlKey || event.metaKey || event.shiftKey || event.altKey) return;
    event.preventDefault();
    selectTab(tab);
    history.replaceState(null, "", tab.hash);
  });
  tab.addEventListener("keydown", event => {
    let next;
    switch (event.key) {
      case " ": next = index; break;
      case "ArrowRight": next = (index + 1) % tabs.length; break;
      case "ArrowLeft": next = (index + tabs.length - 1) % tabs.length; break;
      case "Home": next = 0; break;
      case "End": next = tabs.length - 1; break;
      default: return;
    }
    event.preventDefault();
    tabs[next].focus();
    tabs[next].click();
  });
});
tour.classList.add("enhanced");
selectTab(tabs[0]);
selectFromHash();
window.addEventListener("hashchange", () => {
  selectFromHash();
  const selected = panels.find(panel => "#" + panel.id === location.hash);
  if (selected) selected.scrollIntoView();
});

const feedPreview = tour.querySelector(".feed-preview");
const fullSize = feedPreview.querySelector(".capture");
const feedImage = fullSize.querySelector("img");
const playback = feedPreview.querySelector("button");
const still = fullSize.href;
const reducedMotion = window.matchMedia("(prefers-reduced-motion: reduce)");

function playRecording(play) {
  fullSize.href = play ? playback.dataset.recording : still;
  feedImage.src = fullSize.href;
  playback.textContent = play ? "Stop recording" : "Play recording";
}

playback.hidden = false;
playRecording(!reducedMotion.matches);
playback.addEventListener("click", () => playRecording(fullSize.href === still));
reducedMotion.addEventListener("change", () => playRecording(!reducedMotion.matches));
