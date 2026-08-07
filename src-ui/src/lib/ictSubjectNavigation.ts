export const highlightSubjectElement = (element: HTMLElement) => {
  element.style.transition = "background-color 0.5s ease, box-shadow 0.5s ease";
  element.style.boxShadow = "0 0 0 4px rgba(40, 90, 185, 0.25)";
  element.style.backgroundColor = "rgba(40, 90, 185, 0.05)";
  element.style.borderRadius = "6px";

  setTimeout(() => {
    element.style.boxShadow = "none";
    element.style.backgroundColor = "transparent";
  }, 2000);
};

export const scrollToSubject = (
  side: string,
  groupId: string,
  key: string,
  activeTab: string,
  setActiveTab: (tab: any) => void,
) => {
  const targetTab = side === "revenue" ? "revenue" : "cost";
  const locate = () => {
    const element = document.getElementById(`subject-anchor-${side}-${groupId}-${key}`);
    if (!element) return;
    element.scrollIntoView({ behavior: "smooth", block: "center" });
    highlightSubjectElement(element);
  };

  if (activeTab !== targetTab) {
    setActiveTab(targetTab);
    setTimeout(locate, 150);
  } else {
    locate();
  }
};
