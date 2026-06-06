// In-app navigation history stack.
// Tracks every location change made through wouter so the back button
// navigates within the app rather than reaching into browser history.

const stack: string[] = [];

export function recordPath(path: string) {
  if (stack[stack.length - 1] !== path) {
    stack.push(path);
  }
}

export function goBack(navigate: (to: string) => void, fallback = '/projects') {
  stack.pop(); // remove current page
  const previous = stack[stack.length - 1];
  if (previous) {
    navigate(previous);
  } else {
    navigate(fallback);
  }
}
