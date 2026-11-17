const ALLOWED_NAVIGATION_KEYS = new Set([
  'PageDown',
  'PageUp',
  'ArrowDown',
  'ArrowUp',
  'ArrowLeft',
  'ArrowRight',
  'Home',
  'End',
  'Tab',
  'Escape',
])

const INTERACTIVE_SELECTOR = 'button,[role="button"],[role="menuitem"],[role="tab"],[role="option"],input,select,textarea,label,summary,[contenteditable]'

export function blockBusyEditorInput(event: Event): void {
  const type = event.type
  if (type === 'beforeinput' || type === 'paste' || type === 'cut' || type === 'drop' || type === 'compositionstart' || type === 'dblclick') {
    event.preventDefault()
    event.stopImmediatePropagation?.()
    return
  }

  if (type === 'keydown') {
    const keyEvent = event as KeyboardEvent
    const key = keyEvent.key || ''
    if (ALLOWED_NAVIGATION_KEYS.has(key) && !keyEvent.ctrlKey && !keyEvent.metaKey && !keyEvent.altKey) {
      return
    }
    event.preventDefault()
    event.stopImmediatePropagation?.()
    return
  }

  if (type === 'click' || type === 'pointerdown') {
    const target = event.target as { closest?: (selector: string) => unknown } | null
    if (target && typeof target.closest === 'function' && target.closest(INTERACTIVE_SELECTOR)) {
      event.preventDefault()
    event.stopImmediatePropagation?.()
    }
  }
}
