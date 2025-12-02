/**
 * Prairie Forge Custom Date Picker
 * A modern, dark-themed date picker for the TaiTools add-in
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

// Track active picker for closing on outside click
let activePickerId = null;

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
    
    // Allow manual typing - don't set readonly
    input.type = 'text';
    input.placeholder = 'YYYY-MM-DD or click calendar';
    input.classList.add('pf-datepicker-input');
    
    // Parse initial value
    let selectedDate = input.value ? parseDate(input.value) : null;
    let viewDate = selectedDate ? new Date(selectedDate) : new Date();
    
    // Format display value
    if (selectedDate) {
        input.value = formatDisplayDate(selectedDate);
    }
    
    // Create calendar icon
    const icon = document.createElement('span');
    icon.className = 'pf-datepicker-icon';
    icon.innerHTML = `<svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect width="18" height="18" x="3" y="4" rx="2" ry="2"/><line x1="16" x2="16" y1="2" y2="6"/><line x1="8" x2="8" y1="2" y2="6"/><line x1="3" x2="21" y1="10" y2="10"/></svg>`;
    wrapper.appendChild(icon);
    
    // Create dropdown calendar
    const dropdown = document.createElement('div');
    dropdown.className = 'pf-datepicker-dropdown';
    dropdown.id = `${inputId}-dropdown`;
    wrapper.appendChild(dropdown);
    
    // Render calendar
    function renderCalendar() {
        const year = viewDate.getFullYear();
        const month = viewDate.getMonth();
        
        dropdown.innerHTML = `
            <div class="pf-datepicker-header">
                <button type="button" class="pf-datepicker-nav pf-datepicker-prev-year" title="Previous Year">«</button>
                <button type="button" class="pf-datepicker-nav pf-datepicker-prev" title="Previous Month">‹</button>
                <span class="pf-datepicker-title">${MONTH_NAMES[month]} ${year}</span>
                <button type="button" class="pf-datepicker-nav pf-datepicker-next" title="Next Month">›</button>
                <button type="button" class="pf-datepicker-nav pf-datepicker-next-year" title="Next Year">»</button>
                <button type="button" class="pf-datepicker-nav pf-datepicker-close" title="Close">×</button>
            </div>
            <div class="pf-datepicker-weekdays">
                ${DAY_NAMES.map(d => `<span>${d}</span>`).join('')}
            </div>
            <div class="pf-datepicker-days">
                ${generateDays(year, month, selectedDate)}
            </div>
            <div class="pf-datepicker-footer">
                <button type="button" class="pf-datepicker-today">Today</button>
                <button type="button" class="pf-datepicker-clear">Clear</button>
            </div>
        `;
        
        // Bind navigation - use mousedown to prevent blur events
        dropdown.querySelector('.pf-datepicker-prev-year')?.addEventListener('mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            viewDate.setFullYear(viewDate.getFullYear() - 1);
            renderCalendar();
        });
        
        dropdown.querySelector('.pf-datepicker-prev')?.addEventListener('mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            viewDate.setMonth(viewDate.getMonth() - 1);
            renderCalendar();
        });
        
        dropdown.querySelector('.pf-datepicker-next')?.addEventListener('mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            viewDate.setMonth(viewDate.getMonth() + 1);
            renderCalendar();
        });
        
        dropdown.querySelector('.pf-datepicker-next-year')?.addEventListener('mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            viewDate.setFullYear(viewDate.getFullYear() + 1);
            renderCalendar();
        });
        
        // Close button
        dropdown.querySelector('.pf-datepicker-close')?.addEventListener('mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            closeDropdown();
        });
        
        // Bind day clicks - use mousedown to prevent blur
        dropdown.querySelectorAll('.pf-datepicker-day:not(.disabled)').forEach(dayEl => {
            dayEl.addEventListener('mousedown', (e) => {
                e.preventDefault();
                e.stopPropagation();
                const day = parseInt(dayEl.dataset.day);
                const m = parseInt(dayEl.dataset.month);
                const y = parseInt(dayEl.dataset.year);
                selectDate(new Date(y, m, day));
            });
        });
        
        // Today button
        dropdown.querySelector('.pf-datepicker-today')?.addEventListener('mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            selectDate(new Date());
        });
        
        // Clear button
        dropdown.querySelector('.pf-datepicker-clear')?.addEventListener('mousedown', (e) => {
            e.preventDefault();
            e.stopPropagation();
            selectDate(null);
        });
    }
    
    function generateDays(year, month, selected) {
        const firstDay = new Date(year, month, 1).getDay();
        const daysInMonth = new Date(year, month + 1, 0).getDate();
        const daysInPrevMonth = new Date(year, month, 0).getDate();
        
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        
        let html = '';
        
        // Previous month days
        for (let i = firstDay - 1; i >= 0; i--) {
            const day = daysInPrevMonth - i;
            const prevMonth = month === 0 ? 11 : month - 1;
            const prevYear = month === 0 ? year - 1 : year;
            html += `<span class="pf-datepicker-day other-month" data-day="${day}" data-month="${prevMonth}" data-year="${prevYear}">${day}</span>`;
        }
        
        // Current month days
        for (let day = 1; day <= daysInMonth; day++) {
            const date = new Date(year, month, day);
            const isToday = date.getTime() === today.getTime();
            const isSelected = selected && date.getTime() === selected.getTime();
            
            let classes = 'pf-datepicker-day';
            if (isToday) classes += ' today';
            if (isSelected) classes += ' selected';
            
            // Check min/max constraints
            if (minDate && date < minDate) classes += ' disabled';
            if (maxDate && date > maxDate) classes += ' disabled';
            
            html += `<span class="${classes}" data-day="${day}" data-month="${month}" data-year="${year}">${day}</span>`;
        }
        
        // Next month days - ALWAYS fill to 42 cells (6 rows × 7 days) for consistent height
        const totalCells = 42;
        const currentCells = firstDay + daysInMonth;
        const nextMonthDays = totalCells - currentCells;
        for (let day = 1; day <= nextMonthDays; day++) {
            const nextMonth = month === 11 ? 0 : month + 1;
            const nextYear = month === 11 ? year + 1 : year;
            html += `<span class="pf-datepicker-day other-month" data-day="${day}" data-month="${nextMonth}" data-year="${nextYear}">${day}</span>`;
        }
        
        return html;
    }
    
    function selectDate(date) {
        selectedDate = date;
        if (date) {
            input.value = formatDisplayDate(date);
            input.dataset.value = formatISODate(date);
            viewDate = new Date(date);
        } else {
            input.value = '';
            input.dataset.value = '';
        }
        closeDropdown();
        
        if (onChange) {
            onChange(date ? formatISODate(date) : '');
        }
        
        // Trigger change event
        input.dispatchEvent(new Event('change', { bubbles: true }));
    }
    
    function openDropdown() {
        if (readonly) return;
        
        // Close any other open picker
        if (activePickerId && activePickerId !== inputId) {
            const otherDropdown = document.getElementById(`${activePickerId}-dropdown`);
            otherDropdown?.classList.remove('open');
        }
        
        activePickerId = inputId;
        renderCalendar();
        dropdown.classList.add('open');
        wrapper.classList.add('open');
    }
    
    function closeDropdown() {
        dropdown.classList.remove('open');
        wrapper.classList.remove('open');
        if (activePickerId === inputId) {
            activePickerId = null;
        }
    }
    
    // Handle manual date entry
    input.addEventListener('blur', (e) => {
        // Don't process if clicking within the dropdown
        if (dropdown.classList.contains('open')) return;
        
        const typed = input.value.trim();
        if (!typed) return;
        
        const parsedDate = parseDate(typed);
        if (parsedDate) {
            selectedDate = parsedDate;
            input.value = formatDisplayDate(parsedDate);
            input.dataset.value = formatISODate(parsedDate);
            viewDate = new Date(parsedDate);
            
            if (onChange) {
                onChange(formatISODate(parsedDate));
            }
            input.dispatchEvent(new Event('change', { bubbles: true }));
        }
    });
    
    // Allow typing and Enter to confirm
    input.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
            e.preventDefault();
            const typed = input.value.trim();
            const parsedDate = parseDate(typed);
            if (parsedDate) {
                selectDate(parsedDate);
            }
            closeDropdown();
        }
    });
    
    // Click on input opens dropdown (but allow typing too)
    input.addEventListener('click', (e) => {
        e.stopPropagation();
        // Only toggle if clicking the calendar icon area or input is focused
        if (!dropdown.classList.contains('open')) {
            openDropdown();
        }
    });
    
    // Calendar icon always toggles dropdown
    icon.addEventListener('click', (e) => {
        e.stopPropagation();
        if (dropdown.classList.contains('open')) {
            closeDropdown();
        } else {
            openDropdown();
        }
    });
    
    // Close on outside click - but not when clicking nav buttons
    document.addEventListener('click', (e) => {
        // Check if click is inside the wrapper (includes dropdown)
        if (wrapper.contains(e.target)) {
            return; // Don't close if clicking inside wrapper
        }
        closeDropdown();
    });
    
    // Prevent dropdown from closing when clicking inside it
    dropdown.addEventListener('click', (e) => {
        e.stopPropagation();
    });
    
    // Close on escape
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') {
            closeDropdown();
        }
    });
    
    // Return API
    return {
        getValue: () => selectedDate ? formatISODate(selectedDate) : '',
        setValue: (dateStr) => {
            const date = parseDate(dateStr);
            selectDate(date);
        },
        open: openDropdown,
        close: closeDropdown
    };
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

