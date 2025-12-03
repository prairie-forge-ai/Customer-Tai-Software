/**
 * Prairie Forge Custom Date Picker
 * A modern, modal-style date picker for the TaiTools add-in
 * 
 * © 2025 Prairie Forge LLC
 */

// Month names for display
const MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
];

const MONTH_NAMES_SHORT = [
    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
];

const DAY_NAMES = ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"];

// Single modal instance for all date pickers
let modalEl = null;
let currentPickerState = null;

/**
 * Initialize a date picker on an input element
 * @param {string} inputId - The ID of the input element
 * @param {Object} options - Configuration options
 */
export function initDatePicker(inputId, options = {}) {
    const input = document.getElementById(inputId);
    if (!input) return;
    
    const {
        onChange = null,
        minDate = null,
        maxDate = null,
        readonly = false
    } = options;
    
    // Create wrapper if not exists
    let wrapper = input.closest('.pf-datepicker-wrapper');
    if (!wrapper) {
        wrapper = document.createElement('div');
        wrapper.className = 'pf-datepicker-wrapper';
        input.parentNode.insertBefore(wrapper, input);
        wrapper.appendChild(input);
    }
    
    // Configure input
    input.type = 'text';
    input.placeholder = 'Select date...';
    input.classList.add('pf-datepicker-input');
    input.readOnly = true; // Prevent keyboard on mobile, use modal only
    
    // Parse initial value
    let selectedDate = input.value ? parseDate(input.value) : null;
    
    // Format display value
    if (selectedDate) {
        input.value = formatDisplayDate(selectedDate);
        input.dataset.value = formatISODate(selectedDate);
    }
    
    // Create calendar icon
    let icon = wrapper.querySelector('.pf-datepicker-icon');
    if (!icon) {
        icon = document.createElement('span');
        icon.className = 'pf-datepicker-icon';
        icon.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg>`;
        wrapper.appendChild(icon);
    }
    
    // Picker state
    const pickerState = {
        inputId,
        input,
        selectedDate,
        viewDate: selectedDate ? new Date(selectedDate) : new Date(),
        onChange,
        minDate,
        maxDate
    };
    
    // Open modal on click
    function openModal() {
        if (readonly) return;
        currentPickerState = pickerState;
        showDatePickerModal();
    }
    
    // Click handlers
    input.addEventListener('click', openModal);
    icon.addEventListener('click', openModal);
    
    // Return API
    return {
        getValue: () => pickerState.selectedDate ? formatISODate(pickerState.selectedDate) : '',
        setValue: (dateStr) => {
            const date = parseDate(dateStr);
            pickerState.selectedDate = date;
            pickerState.viewDate = date ? new Date(date) : new Date();
            if (date) {
                input.value = formatDisplayDate(date);
                input.dataset.value = formatISODate(date);
            } else {
                input.value = '';
                input.dataset.value = '';
            }
        },
        open: openModal,
        close: closeDatePickerModal
    };
}

/**
 * Show the date picker modal
 */
function showDatePickerModal() {
    if (!currentPickerState) return;
    
    // Create modal if doesn't exist
    if (!modalEl) {
        modalEl = document.createElement('div');
        modalEl.className = 'pf-datepicker-modal';
        modalEl.id = 'pf-datepicker-modal';
        document.body.appendChild(modalEl);
    }
    
    renderModal();
    
    // Show with animation
    requestAnimationFrame(() => {
        modalEl.classList.add('is-open');
    });
    
    // Close on escape
    document.addEventListener('keydown', handleEscapeKey);
}

/**
 * Close the date picker modal
 */
function closeDatePickerModal() {
    if (modalEl) {
        modalEl.classList.remove('is-open');
    }
    document.removeEventListener('keydown', handleEscapeKey);
    currentPickerState = null;
}

function handleEscapeKey(e) {
    if (e.key === 'Escape') {
        closeDatePickerModal();
    }
}

/**
 * Render the modal content
 */
function renderModal() {
    if (!modalEl || !currentPickerState) return;
    
    const { viewDate, selectedDate, minDate, maxDate } = currentPickerState;
    const year = viewDate.getFullYear();
    const month = viewDate.getMonth();
    
    modalEl.innerHTML = `
        <div class="pf-datepicker-backdrop"></div>
        <div class="pf-datepicker-container">
            <div class="pf-datepicker-header">
                <div class="pf-datepicker-nav-group">
                    <button type="button" class="pf-datepicker-nav" data-action="prev-year" title="Previous Year">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="11 17 6 12 11 7"/><polyline points="18 17 13 12 18 7"/></svg>
                    </button>
                    <button type="button" class="pf-datepicker-nav" data-action="prev-month" title="Previous Month">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="15 18 9 12 15 6"/></svg>
                    </button>
                </div>
                <span class="pf-datepicker-title">${MONTH_NAMES[month]} ${year}</span>
                <div class="pf-datepicker-nav-group">
                    <button type="button" class="pf-datepicker-nav" data-action="next-month" title="Next Month">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="9 18 15 12 9 6"/></svg>
                    </button>
                    <button type="button" class="pf-datepicker-nav" data-action="next-year" title="Next Year">
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="13 17 18 12 13 7"/><polyline points="6 17 11 12 6 7"/></svg>
                    </button>
                </div>
            </div>
            <div class="pf-datepicker-weekdays">
                ${DAY_NAMES.map(d => `<span>${d}</span>`).join('')}
            </div>
            <div class="pf-datepicker-days">
                ${generateDays(year, month, selectedDate, minDate, maxDate)}
            </div>
            <div class="pf-datepicker-footer">
                <button type="button" class="pf-datepicker-btn pf-datepicker-today" data-action="today">Today</button>
                <button type="button" class="pf-datepicker-btn pf-datepicker-clear" data-action="clear">Clear</button>
            </div>
        </div>
    `;
    
    // Bind all events
    bindModalEvents();
}

/**
 * Bind modal event handlers
 */
function bindModalEvents() {
    if (!modalEl) return;
    
    // Backdrop click to close
    modalEl.querySelector('.pf-datepicker-backdrop')?.addEventListener('click', closeDatePickerModal);
    
    // Navigation buttons
    modalEl.querySelectorAll('.pf-datepicker-nav').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            const action = btn.dataset.action;
            handleNavigation(action);
        });
    });
    
    // Day clicks
    modalEl.querySelectorAll('.pf-datepicker-day:not(.disabled)').forEach(dayEl => {
        dayEl.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            const day = parseInt(dayEl.dataset.day);
            const month = parseInt(dayEl.dataset.month);
            const year = parseInt(dayEl.dataset.year);
            selectDate(new Date(year, month, day));
        });
    });
    
    // Footer buttons
    modalEl.querySelectorAll('.pf-datepicker-btn').forEach(btn => {
        btn.addEventListener('click', (e) => {
            e.preventDefault();
            e.stopPropagation();
            const action = btn.dataset.action;
            if (action === 'today') {
                selectDate(new Date());
            } else if (action === 'clear') {
                selectDate(null);
            }
        });
    });
}

/**
 * Handle navigation actions
 */
function handleNavigation(action) {
    if (!currentPickerState) return;
    
    const viewDate = currentPickerState.viewDate;
    
    switch (action) {
        case 'prev-year':
            viewDate.setFullYear(viewDate.getFullYear() - 1);
            break;
        case 'prev-month':
            viewDate.setMonth(viewDate.getMonth() - 1);
            break;
        case 'next-month':
            viewDate.setMonth(viewDate.getMonth() + 1);
            break;
        case 'next-year':
            viewDate.setFullYear(viewDate.getFullYear() + 1);
            break;
    }
    
    renderModal();
}

/**
 * Select a date and close modal
 */
function selectDate(date) {
    if (!currentPickerState) return;
    
    const { input, onChange } = currentPickerState;
    
    currentPickerState.selectedDate = date;
    
    if (date) {
        input.value = formatDisplayDate(date);
        input.dataset.value = formatISODate(date);
        currentPickerState.viewDate = new Date(date);
    } else {
        input.value = '';
        input.dataset.value = '';
    }
    
    // Trigger callbacks
    if (onChange) {
        onChange(date ? formatISODate(date) : '');
    }
    input.dispatchEvent(new Event('change', { bubbles: true }));
    
    // Close modal
    closeDatePickerModal();
}

/**
 * Generate day cells
 */
function generateDays(year, month, selected, minDate, maxDate) {
    const firstDay = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const daysInPrevMonth = new Date(year, month, 0).getDate();
    
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    
    if (selected) {
        selected = new Date(selected);
        selected.setHours(0, 0, 0, 0);
    }
    
    let html = '';
    
    // Previous month days
    for (let i = firstDay - 1; i >= 0; i--) {
        const day = daysInPrevMonth - i;
        const prevMonth = month === 0 ? 11 : month - 1;
        const prevYear = month === 0 ? year - 1 : year;
        html += `<button type="button" class="pf-datepicker-day other-month" data-day="${day}" data-month="${prevMonth}" data-year="${prevYear}">${day}</button>`;
    }
    
    // Current month days
    for (let day = 1; day <= daysInMonth; day++) {
        const date = new Date(year, month, day);
        date.setHours(0, 0, 0, 0);
        const isToday = date.getTime() === today.getTime();
        const isSelected = selected && date.getTime() === selected.getTime();
        
        let classes = 'pf-datepicker-day';
        if (isToday) classes += ' today';
        if (isSelected) classes += ' selected';
        
        // Check min/max constraints
        let disabled = false;
        if (minDate && date < minDate) disabled = true;
        if (maxDate && date > maxDate) disabled = true;
        if (disabled) classes += ' disabled';
        
        html += `<button type="button" class="${classes}" data-day="${day}" data-month="${month}" data-year="${year}" ${disabled ? 'disabled' : ''}>${day}</button>`;
    }
    
    // Next month days - fill to 42 cells (6 rows)
    const totalCells = 42;
    const currentCells = firstDay + daysInMonth;
    const nextMonthDays = totalCells - currentCells;
    for (let day = 1; day <= nextMonthDays; day++) {
        const nextMonth = month === 11 ? 0 : month + 1;
        const nextYear = month === 11 ? year + 1 : year;
        html += `<button type="button" class="pf-datepicker-day other-month" data-day="${day}" data-month="${nextMonth}" data-year="${nextYear}">${day}</button>`;
    }
    
    return html;
}

/**
 * Parse various date formats to Date object
 */
function parseDate(str) {
    if (!str) return null;
    
    // Handle ISO format (YYYY-MM-DD)
    if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
        const [y, m, d] = str.split('-').map(Number);
        return new Date(y, m - 1, d);
    }
    
    // Handle display format (MMM D, YYYY)
    const displayMatch = str.match(/^(\w+)\s+(\d+),\s+(\d{4})$/);
    if (displayMatch) {
        const monthIdx = MONTH_NAMES_SHORT.findIndex(m => 
            m.toLowerCase() === displayMatch[1].toLowerCase().substring(0, 3)
        );
        if (monthIdx >= 0) {
            return new Date(parseInt(displayMatch[3]), monthIdx, parseInt(displayMatch[2]));
        }
    }
    
    // Handle MM/DD/YYYY
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)) {
        const [m, d, y] = str.split('/').map(Number);
        return new Date(y, m - 1, d);
    }
    
    // Fallback to Date.parse
    const date = new Date(str);
    return isNaN(date.getTime()) ? null : date;
}

/**
 * Format date for display (Nov 30, 2025)
 */
function formatDisplayDate(date) {
    if (!date) return '';
    return `${MONTH_NAMES_SHORT[date.getMonth()]} ${date.getDate()}, ${date.getFullYear()}`;
}

/**
 * Format date as ISO (2025-11-30)
 */
function formatISODate(date) {
    if (!date) return '';
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
}

// Export utilities
export { parseDate, formatDisplayDate, formatISODate };
