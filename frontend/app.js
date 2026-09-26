/**
 * Follow-Up CRM Frontend Application
 * Interacts with Node.js Express REST API (/api/search, /api/save)
 */

document.addEventListener('DOMContentLoaded', () => {
  // DOM Elements
  const form = document.getElementById('followupForm');
  const mobileInput = document.getElementById('mobileInput');
  const nameInput = document.getElementById('nameInput');
  const suggestionsList = document.getElementById('suggestionsList');
  const searchSpinner = document.getElementById('searchSpinner');
  const previousRatingContainer = document.getElementById('previousRatingContainer');
  const previousRatingStars = document.getElementById('previousRatingStars');
  const previousRemarkContainer = document.getElementById('previousRemarkContainer');
  const previousRemarkText = document.getElementById('previousRemarkText');
  const followupDateInput = document.getElementById('followupDateInput');
  const starIcons = document.querySelectorAll('.star-icon');
  const ratingInput = document.getElementById('ratingInput');
  const ratingText = document.getElementById('ratingText');
  const ratingError = document.getElementById('ratingError');
  const remarkInput = document.getElementById('remarkInput');
  const anyIssuesCheckbox = document.getElementById('anyIssuesCheckbox');
  const reasonContainer = document.getElementById('reasonContainer');
  const reasonInput = document.getElementById('reasonInput');
  const reasonError = document.getElementById('reasonError');
  const submitBtn = document.getElementById('submitBtn');
  const btnSpinner = document.getElementById('btnSpinner');
  const btnText = document.getElementById('btnText');
  const resetBtn = document.getElementById('resetBtn');
  const alertContainer = document.getElementById('alertContainer');

  // Search state
  let searchDebounceTimer = null;
  let activeAbortController = null;
  let currentSearchSeq = 0;
  let searchResultsCache = [];

  const ratingDescriptions = {
    1: '1 Star - Poor',
    2: '2 Stars - Fair',
    3: '3 Stars - Good',
    4: '4 Stars - Very Good',
    5: '5 Stars - Excellent'
  };

  // -------------------------------------------------------------
  // 1. Initialize Date Constraints (Today to Today + 30 Days)
  // -------------------------------------------------------------
  function initializeDateConstraints() {
    const today = new Date();
    const yyyy = today.getFullYear();
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const dd = String(today.getDate()).padStart(2, '0');
    const minDateStr = `${yyyy}-${mm}-${dd}`;

    const maxDate = new Date(today);
    maxDate.setDate(maxDate.getDate() + 30);
    const maxYyyy = maxDate.getFullYear();
    const maxMm = String(maxDate.getMonth() + 1).padStart(2, '0');
    const maxDd = String(maxDate.getDate()).padStart(2, '0');
    const maxDateStr = `${maxYyyy}-${maxMm}-${maxDd}`;

    followupDateInput.min = minDateStr;
    followupDateInput.max = maxDateStr;
    followupDateInput.value = minDateStr;
  }

  initializeDateConstraints();

  // -------------------------------------------------------------
  // 2. Alert Notifications
  // -------------------------------------------------------------
  function showAlert(message, type = 'success') {
    const alertHtml = `
      <div class="alert alert-${type} alert-dismissible fade show shadow-sm" role="alert">
        <i class="bi bi-${type === 'success' ? 'check-circle-fill' : type === 'info' ? 'info-circle-fill' : 'exclamation-triangle-fill'} me-2"></i>
        ${message}
        <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close" onclick="this.parentElement.remove()"></button>
      </div>
    `;
    alertContainer.innerHTML = alertHtml;

    if (type === 'success' || type === 'info') {
      setTimeout(() => {
        const currentAlert = alertContainer.querySelector('.alert');
        if (currentAlert) currentAlert.remove();
      }, 6000);
    }
  }

  function clearAlerts() {
    alertContainer.innerHTML = '';
  }

  // -------------------------------------------------------------
  // 3. Customer Autocomplete Search (Mobile or Name, min 3 chars)
  // -------------------------------------------------------------
  mobileInput.addEventListener('input', (e) => {
    const query = e.target.value.trim();

    // Reset customer selection state when user edits search box
    nameInput.readOnly = false;
    previousRatingContainer.classList.add('d-none');
    previousRemarkContainer.classList.add('d-none');

    clearTimeout(searchDebounceTimer);

    if (query.length < 3) {
      hideSuggestions();
      searchSpinner.classList.add('d-none');
      if (activeAbortController) {
        activeAbortController.abort();
      }
      return;
    }

    // Debounce search ~150ms
    searchSpinner.classList.remove('d-none');
    searchDebounceTimer = setTimeout(() => {
      performCustomerSearch(query);
    }, 150);
  });

  async function performCustomerSearch(query) {
    if (activeAbortController) {
      activeAbortController.abort();
    }
    activeAbortController = new AbortController();
    const thisSeq = ++currentSearchSeq;

    try {
      const response = await fetch(`/api/search?mobile=${encodeURIComponent(query)}`, {
        signal: activeAbortController.signal
      });

      if (!response.ok) {
        throw new Error('Search failed with status ' + response.status);
      }

      const results = await response.json();

      // Guard against stale responses
      if (thisSeq !== currentSearchSeq) return;

      searchResultsCache = results || [];
      renderSuggestions(searchResultsCache);
    } catch (err) {
      if (err.name === 'AbortError') return;
      console.error('Search error:', err);
      suggestionsList.innerHTML = `<li class="dropdown-item text-danger small">Error fetching customers</li>`;
      suggestionsList.style.display = 'block';
    } finally {
      if (thisSeq === currentSearchSeq) {
        searchSpinner.classList.add('d-none');
      }
    }
  }

  function renderSuggestions(customers) {
    suggestionsList.innerHTML = '';

    if (!customers || customers.length === 0) {
      const li = document.createElement('li');
      li.className = 'dropdown-item text-muted disabled fst-italic';
      li.textContent = 'No customer found.';
      suggestionsList.appendChild(li);
      suggestionsList.style.display = 'block';
      return;
    }

    customers.slice(0, 20).forEach((cust, index) => {
      const li = document.createElement('li');
      li.className = 'dropdown-item d-flex justify-content-between align-items-center';
      li.setAttribute('data-index', index);

      const ratingBadge = cust.rating ? `<span class="badge bg-warning text-dark ms-2">★ ${cust.rating}</span>` : '';

      li.innerHTML = `
        <div>
          <span class="customer-name d-block text-dark">${cust.name || 'Unknown Name'}</span>
          <span class="customer-mobile text-muted small"><i class="bi bi-telephone me-1"></i>${cust.mobile}</span>
        </div>
        <div>${ratingBadge}</div>
      `;

      li.addEventListener('click', () => selectCustomer(cust));
      suggestionsList.appendChild(li);
    });

    suggestionsList.style.display = 'block';
  }

  function hideSuggestions() {
    suggestionsList.style.display = 'none';
    suggestionsList.innerHTML = '';
  }

  // Hide suggestions when clicking outside
  document.addEventListener('click', (e) => {
    if (!mobileInput.contains(e.target) && !suggestionsList.contains(e.target)) {
      hideSuggestions();
    }
  });

  // -------------------------------------------------------------
  // 4. Customer Selection & Auto-fill
  // -------------------------------------------------------------
  function selectCustomer(cust) {
    mobileInput.value = cust.mobile || '';
    nameInput.value = cust.name || '';
    nameInput.readOnly = true;

    // Previous Rating
    if (cust.rating && !isNaN(parseInt(cust.rating, 10))) {
      const rateNum = Math.min(5, Math.max(1, parseInt(cust.rating, 10)));
      const solidStars = '★'.repeat(rateNum);
      const emptyStars = '☆'.repeat(5 - rateNum);
      previousRatingStars.innerHTML = `<span class="me-2">${solidStars}${emptyStars}</span> <span class="badge bg-secondary fs-6 align-middle">${rateNum} / 5</span>`;
      previousRatingContainer.classList.remove('d-none');
    } else {
      previousRatingContainer.classList.add('d-none');
      previousRatingStars.innerHTML = '';
    }

    // Previous Remark
    if (cust.remark && cust.remark.trim() !== '') {
      previousRemarkText.value = cust.remark.trim();
      previousRemarkContainer.classList.remove('d-none');
    } else {
      previousRemarkContainer.classList.add('d-none');
      previousRemarkText.value = '';
    }

    // Reset current follow-up rating & issues
    resetCurrentRating();
    anyIssuesCheckbox.checked = false;
    reasonInput.value = '';
    reasonContainer.classList.add('d-none');
    remarkInput.value = '';

    hideSuggestions();
  }

  // -------------------------------------------------------------
  // 5. Star Rating UI
  // -------------------------------------------------------------
  function updateStars(val) {
    starIcons.forEach(icon => {
      const iconVal = parseInt(icon.getAttribute('data-value'), 10);
      const bi = icon.querySelector('i');
      if (iconVal <= val) {
        icon.classList.add('active');
        bi.className = 'bi bi-star-fill text-warning';
      } else {
        icon.classList.remove('active');
        bi.className = 'bi bi-star text-muted';
      }
    });

    if (val > 0) {
      ratingText.textContent = ratingDescriptions[val] || `${val} Stars`;
      ratingError.classList.add('d-none');
    } else {
      ratingText.textContent = '';
    }
  }

  function resetCurrentRating() {
    ratingInput.value = '';
    updateStars(0);
    ratingError.classList.add('d-none');
  }

  starIcons.forEach(icon => {
    // Hover effect
    icon.addEventListener('mouseenter', () => {
      const hoverVal = parseInt(icon.getAttribute('data-value'), 10);
      starIcons.forEach(s => {
        const sVal = parseInt(s.getAttribute('data-value'), 10);
        const bi = s.querySelector('i');
        if (sVal <= hoverVal) {
          bi.className = 'bi bi-star-fill text-warning';
        } else {
          bi.className = 'bi bi-star text-muted';
        }
      });
    });

    // Restore selected on mouse leave
    icon.addEventListener('mouseleave', () => {
      const selectedVal = parseInt(ratingInput.value, 10) || 0;
      updateStars(selectedVal);
    });

    // Click to select
    icon.addEventListener('click', () => {
      const selectedVal = parseInt(icon.getAttribute('data-value'), 10);
      ratingInput.value = selectedVal;
      updateStars(selectedVal);
    });
  });

  // -------------------------------------------------------------
  // 6. Any Issues Checkbox & Reason Toggle
  // -------------------------------------------------------------
  anyIssuesCheckbox.addEventListener('change', () => {
    if (anyIssuesCheckbox.checked) {
      reasonContainer.classList.remove('d-none');
      reasonInput.required = true;
      reasonInput.focus();
    } else {
      reasonContainer.classList.add('d-none');
      reasonInput.required = false;
      reasonInput.value = '';
      reasonError.classList.add('d-none');
      reasonInput.classList.remove('is-invalid');
    }
  });

  reasonInput.addEventListener('input', () => {
    if (reasonInput.value.trim()) {
      reasonError.classList.add('d-none');
      reasonInput.classList.remove('is-invalid');
    }
  });

  // -------------------------------------------------------------
  // 7. Form Reset
  // -------------------------------------------------------------
  function resetForm() {
    form.reset();
    nameInput.readOnly = false;
    previousRatingContainer.classList.add('d-none');
    previousRatingStars.innerHTML = '';
    previousRemarkContainer.classList.add('d-none');
    previousRemarkText.value = '';
    resetCurrentRating();
    reasonContainer.classList.add('d-none');
    reasonError.classList.add('d-none');
    reasonInput.classList.remove('is-invalid');
    hideSuggestions();
    initializeDateConstraints();
    mobileInput.focus();
  }

  resetBtn.addEventListener('click', () => {
    resetForm();
    clearAlerts();
  });

  // -------------------------------------------------------------
  // 8. Form Validation & Submission
  // -------------------------------------------------------------
  form.addEventListener('submit', async (e) => {
    e.preventDefault();
    clearAlerts();

    const mobile = mobileInput.value.trim();
    const name = nameInput.value.trim();
    const followup = followupDateInput.value;
    const rating = ratingInput.value;
    const remark = remarkInput.value.trim();
    const anyIssues = anyIssuesCheckbox.checked ? 'Yes' : 'No';
    const reason = reasonInput.value.trim();

    // Client-side validations
    let isValid = true;

    // Mobile validation
    const cleanMobile = mobile.replace(/\D/g, '');
    if (!cleanMobile || cleanMobile.length !== 10) {
      showAlert('Please select a customer from the search list or enter a valid 10-digit mobile number.', 'danger');
      mobileInput.focus();
      return;
    }

    // Name validation
    if (!name) {
      showAlert('Customer Name is required.', 'danger');
      nameInput.focus();
      return;
    }

    // Date validation
    if (!followup) {
      showAlert('Please select a follow-up date.', 'danger');
      followupDateInput.focus();
      return;
    }

    const todayDate = new Date();
    todayDate.setHours(0, 0, 0, 0);
    const selectedDate = new Date(followup + 'T00:00:00');
    const maxAllowedDate = new Date(todayDate);
    maxAllowedDate.setDate(maxAllowedDate.getDate() + 30);
    maxAllowedDate.setHours(23, 59, 59, 999);

    if (selectedDate < todayDate) {
      showAlert('Follow-up date cannot be in the past.', 'danger');
      followupDateInput.focus();
      return;
    }

    if (selectedDate > maxAllowedDate) {
      showAlert('Follow-up date cannot be more than 30 days ahead.', 'danger');
      followupDateInput.focus();
      return;
    }

    // Rating validation
    if (!rating || parseInt(rating, 10) < 1 || parseInt(rating, 10) > 5) {
      ratingError.classList.remove('d-none');
      showAlert('Please select a rating between 1 and 5 stars.', 'danger');
      return;
    }

    // Reason validation when Any Issues is checked
    if (anyIssues === 'Yes' && !reason) {
      reasonError.classList.remove('d-none');
      reasonInput.classList.add('is-invalid');
      showAlert('Reason is required when "Any Issues" is checked.', 'danger');
      reasonInput.focus();
      return;
    }

    // Generate unique submission ID (UUID v4)
    const submissionId = (typeof crypto !== 'undefined' && crypto.randomUUID)
      ? crypto.randomUUID()
      : 'sub_' + Date.now() + '_' + Math.random().toString(36).substring(2, 10);

    // Format follow-up date as DD/MM/YYYY (e.g. 26/09/2026)
    let formattedFollowup = followup;
    if (followup && followup.includes('-')) {
      const fParts = followup.split('-');
      if (fParts.length === 3) {
        formattedFollowup = `${fParts[2]}/${fParts[1]}/${fParts[0]}`;
      }
    }

    const payload = {
      mobile,
      name,
      followup: formattedFollowup,
      rating: parseInt(rating, 10),
      remark,
      anyIssues,
      reason: anyIssues === 'Yes' ? reason : '',
      submissionId
    };

    // UI Loading state
    submitBtn.disabled = true;
    btnSpinner.classList.remove('d-none');
    btnText.textContent = 'Saving data...';

    try {
      const response = await fetch('/api/save', {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(payload)
      });

      const data = await response.json();

      if (response.status === 409 || (data.message && data.message.toLowerCase().includes('already saved'))) {
        showAlert('This follow-up entry has already been saved.', 'info');
        resetForm();
        return;
      }

      if (!response.ok) {
        throw new Error(data.message || 'Server error saving follow-up data.');
      }

      showAlert(`Follow-up record for <strong>${name}</strong> (${mobile}) saved successfully!`, 'success');
      resetForm();
    } catch (err) {
      console.error('Save error:', err);
      showAlert(err.message || 'Failed to connect to the server. Please check your network and try again.', 'danger');
    } finally {
      submitBtn.disabled = false;
      btnSpinner.classList.add('d-none');
      btnText.innerHTML = '<i class="bi bi-send-fill me-1"></i> Submit Follow-Up';
    }
  });
});
