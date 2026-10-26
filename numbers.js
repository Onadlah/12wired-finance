// Group digits for display without rounding or changing the stored value.
export function grouped(value) {
  const raw = String(value ?? '').replaceAll(',', '');
  if (!/^-?\d*(?:\.\d*)?$/.test(raw)) return String(value ?? '');
  const [whole, fraction] = raw.split('.');
  return whole.replace(/\B(?=(\d{3})+(?!\d))/g, ',') +
    (fraction === undefined ? '' : '.' + fraction);
}

export function numericValue(value) {
  const raw = String(value ?? '').replaceAll(',', '').trim();
  if (!/^-?(?:\d+(?:\.\d*)?|\.\d+)$/.test(raw) || !Number.isFinite(Number(raw))) {
    throw Error('Enter a valid number, for example 1,250,000.50.');
  }
  return Number(raw);
}

function validate(input) {
  input.setCustomValidity('');
  if (input.value === '') return;
  try {
    const value = numericValue(input.value);
    if (input.min !== '' && value < Number(input.min)) throw Error(`Minimum: ${grouped(input.min)}`);
    if (input.max !== '' && value > Number(input.max)) throw Error(`Maximum: ${grouped(input.max)}`);
    const step = input.step || '1';
    const offset = (value - Number(input.min || 0)) / Number(step);
    if (step !== 'any' && Math.abs(offset - Math.round(offset)) > Math.max(1e-7, Math.abs(offset) * Number.EPSILON * 4)) {
      throw Error(`Use increments of ${step}.`);
    }
  } catch (error) {
    input.setCustomValidity(error.message);
  }
}

export function enhanceNumbers(root = document) {
  root.querySelectorAll('input[type="number"]').forEach(input => {
    input.type = 'text';
    input.inputMode = 'decimal';
    input.dataset.numeric = '';
    input.value = grouped(input.value);
    input.addEventListener('input', () => {
      // Keep the caret beside the same digit when grouping inserts commas.
      const position = input.selectionStart ?? input.value.length;
      const digitsBefore = input.value.slice(0, position).replaceAll(',', '').length;
      input.value = grouped(input.value);
      let caret = 0, seen = 0;
      while (caret < input.value.length && seen < digitsBefore) {
        if (input.value[caret] !== ',') seen++;
        caret++;
      }
      input.setSelectionRange(caret, caret);
      validate(input);
    });
    input.addEventListener('blur', () => validate(input));
    validate(input);
  });
}
