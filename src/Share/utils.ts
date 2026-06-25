/**
 * Utility to force scrolling modern page components and SharePoint scroll wrappers to the top.
 */
export const forceScrollToTop = () => {
  // 1. Reset standard window/document scrolling
  window.scrollTo(0, 0);
  document.documentElement.scrollTo(0, 0);
  document.body.scrollTo(0, 0);
  document.documentElement.scrollTop = 0;
  document.body.scrollTop = 0;

  // 2. Select common SharePoint framework/SPFx scroll containers
  const selectors = [
    '[data-automation-id="contentScrollRegion"]',
    '#workbenchPageContent',
    '.ms-ScrollablePane--contentContainer',
    '.CanvasZone',
    'main',
    '.pageContainer'
  ];

  selectors.forEach(selector => {
    try {
      const elements = document.querySelectorAll(selector);
      elements.forEach((el: any) => {
        if (el) {
          el.scrollTop = 0;
          if (typeof el.scrollTo === 'function') {
            el.scrollTo(0, 0);
          }
        }
      });
    } catch (e) {
      console.warn("Failed to reset scroll for selector:", selector, e);
    }
  });
};
