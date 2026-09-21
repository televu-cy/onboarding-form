// Shared splitforms submission helpers used by index.html and index2.html.
window.OnboardingSubmit = (() => {
  // splitforms accepts at most 100 fields per submission; each user takes 8.
  const MAX_USERS = 10;

  function buildSubmissionId() {
    const date = new Date().toISOString().slice(0, 10);
    const suffix = Math.random().toString(36).slice(2, 7).toUpperCase();
    return `OF-${date}-${suffix}`;
  }

  function joinList(value) {
    return Array.isArray(value) ? value.join(', ') : '';
  }

  function failureMessage(status, data) {
    if (status === 429) {
      return 'Too many submissions right now. Please wait a minute and try again.';
    }
    const detail = data && data.message ? ` (${data.message})` : '';
    return `We couldn't save your request${detail}. Please try again, or contact TeleVU if this keeps happening.`;
  }

  // Resolves only when splitforms confirms the submission was stored; throws otherwise.
  async function send(fields) {
    const config = window.APP_CONFIG || {};
    if (!config.SPLITFORMS_ENDPOINT || !config.SPLITFORMS_ACCESS_KEY) {
      throw new Error('This form is not configured yet. Please contact TeleVU.');
    }

    let response;
    try {
      response = await fetch(config.SPLITFORMS_ENDPOINT, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', Accept: 'application/json' },
        body: JSON.stringify({ access_key: config.SPLITFORMS_ACCESS_KEY, ...fields }),
      });
    } catch (error) {
      throw new Error("We couldn't reach our server. Check your internet connection and try again.");
    }

    let data = null;
    try {
      data = await response.json();
    } catch (error) {
      // Non-JSON reply; treated as a failure below.
    }

    if (!response.ok || !data || data.success !== true) {
      throw new Error(failureMessage(response.status, data));
    }
    return data;
  }

  return { MAX_USERS, buildSubmissionId, joinList, send };
})();
