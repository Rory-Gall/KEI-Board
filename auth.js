/* Microsoft sign-in for the KEI board.
 *
 * Written straight against Microsoft's endpoints (authorization code + PKCE),
 * so the page loads no third-party library and there is no secret anywhere.
 * The board is a public client: the client id is public, and what a person can
 * see is decided by their own SharePoint permissions, not by this file.
 *
 *   AUTH.start()    - call once on load. Finishes a sign-in redirect if we are
 *                     coming back from one. Resolves to {signedIn, name, email}.
 *   AUTH.signIn()   - send the person to Microsoft (a redirect, not a popup:
 *                     popups get blocked, and this board is a full page anyway).
 *   AUTH.token()    - a valid access token, refreshed silently. Rejects with
 *                     AUTH.EXPIRED when the person must sign in again.
 *   AUTH.signOut()  - forget the tokens on this device.
 *
 * Tokens live in sessionStorage: they die with the tab, which is the right
 * default on a shared office PC. The refresh token keeps a working day going
 * without re-prompting.
 */
var AUTH = (function () {
  'use strict';

  var AUTHORITY = 'https://login.microsoftonline.com/' + CONFIG.tenantId + '/oauth2/v2.0';
  var SCOPES = 'https://graph.microsoft.com/Sites.ReadWrite.All https://graph.microsoft.com/User.Read offline_access openid profile';
  var KEY = 'kei_board_auth';
  var VERIFIER_KEY = 'kei_board_pkce';
  var RETURN_KEY = 'kei_board_return';
  var EXPIRED = 'signed-out';

  var state = null;   /* {access_token, refresh_token, expires_at, name, email} */

  function redirectUri() {
    /* Must match the app registration exactly, including the trailing slash. */
    return location.origin + location.pathname.replace(/[^/]*$/, '');
  }
  function load() {
    try { return JSON.parse(sessionStorage.getItem(KEY) || 'null'); } catch (e) { return null; }
  }
  function save(s) {
    state = s;
    try { sessionStorage.setItem(KEY, JSON.stringify(s)); } catch (e) {}
  }
  function clear() {
    state = null;
    try { sessionStorage.removeItem(KEY); sessionStorage.removeItem(VERIFIER_KEY); } catch (e) {}
  }

  function b64url(bytes) {
    var s = '';
    new Uint8Array(bytes).forEach(function (b) { s += String.fromCharCode(b); });
    return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
  }
  function randomVerifier() {
    var a = new Uint8Array(32);
    crypto.getRandomValues(a);
    return b64url(a);
  }
  function challenge(verifier) {
    return crypto.subtle.digest('SHA-256', new TextEncoder().encode(verifier)).then(b64url);
  }

  function form(url, data) {
    var body = Object.keys(data).map(function (k) {
      return encodeURIComponent(k) + '=' + encodeURIComponent(data[k]);
    }).join('&');
    return fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: body
    }).then(function (r) { return r.json(); });
  }

  /* The signed-in person's name and address, read out of the id token. Saves a
     Graph call on every load, and it is the same information. */
  function identity(tok) {
    var out = { name: '', email: '' };
    if (!tok.id_token) return out;
    try {
      var p = tok.id_token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/');
      var c = JSON.parse(decodeURIComponent(escape(atob(p + '==='.slice((p.length + 3) % 4)))));
      out.name = c.name || '';
      out.email = (c.preferred_username || c.upn || '').toLowerCase();
    } catch (e) {}
    return out;
  }

  function keep(tok) {
    if (!tok.access_token) return null;
    var who = identity(tok);
    var s = {
      access_token: tok.access_token,
      refresh_token: tok.refresh_token || (state && state.refresh_token) || '',
      expires_at: Date.now() + (Number(tok.expires_in || 3600) * 1000),
      name: who.name || (state && state.name) || '',
      email: who.email || (state && state.email) || ''
    };
    save(s);
    return s;
  }

  function signIn() {
    var verifier = randomVerifier();
    try {
      sessionStorage.setItem(VERIFIER_KEY, verifier);
      /* Microsoft sends people back to the registered address, which carries no
         query of ours - so remember where they actually were and put them back
         there afterwards. */
      sessionStorage.setItem(RETURN_KEY, location.search + location.hash);
    } catch (e) {}
    return challenge(verifier).then(function (ch) {
      var q = {
        client_id: CONFIG.clientId, response_type: 'code', redirect_uri: redirectUri(),
        response_mode: 'query', scope: SCOPES, code_challenge: ch, code_challenge_method: 'S256',
        prompt: 'select_account'
      };
      location.href = AUTHORITY + '/authorize?' + Object.keys(q).map(function (k) {
        return k + '=' + encodeURIComponent(q[k]);
      }).join('&');
    });
  }

  /* Coming back from Microsoft: swap the one-time code for tokens, then take
     the code out of the address bar so a refresh cannot replay it. */
  function finishRedirect() {
    var p = new URLSearchParams(location.search);
    var code = p.get('code'), err = p.get('error');
    if (err) {
      history.replaceState({}, '', redirectUri());
      return Promise.reject(new Error(p.get('error_description') || err));
    }
    if (!code) return Promise.resolve(null);
    var verifier = '';
    try { verifier = sessionStorage.getItem(VERIFIER_KEY) || ''; } catch (e) {}
    return form(AUTHORITY + '/token', {
      client_id: CONFIG.clientId, grant_type: 'authorization_code', code: code,
      redirect_uri: redirectUri(), code_verifier: verifier, scope: SCOPES
    }).then(function (tok) {
      var back = '';
      try {
        back = sessionStorage.getItem(RETURN_KEY) || '';
        sessionStorage.removeItem(VERIFIER_KEY);
        sessionStorage.removeItem(RETURN_KEY);
      } catch (e) {}
      history.replaceState({}, '', redirectUri() + back);
      if (!tok.access_token) throw new Error(tok.error_description || tok.error || 'sign-in failed');
      return keep(tok);
    });
  }

  var refreshing = null;
  function token() {
    if (!state) state = load();
    if (!state) return Promise.reject(new Error(EXPIRED));
    if (state.expires_at - 120000 > Date.now()) return Promise.resolve(state.access_token);
    if (refreshing) return refreshing;
    if (!state.refresh_token) { clear(); return Promise.reject(new Error(EXPIRED)); }
    refreshing = form(AUTHORITY + '/token', {
      client_id: CONFIG.clientId, grant_type: 'refresh_token',
      refresh_token: state.refresh_token, scope: SCOPES
    }).then(function (tok) {
      refreshing = null;
      if (!tok.access_token) { clear(); throw new Error(EXPIRED); }
      return keep(tok).access_token;
    }, function (e) {
      refreshing = null;
      throw e;   /* network trouble is NOT a sign-out: let the caller retry */
    });
    return refreshing;
  }

  function start() {
    state = load();
    return finishRedirect().then(function (fresh) {
      var s = fresh || state;
      return s
        ? { signedIn: true, name: s.name, email: s.email }
        : { signedIn: false, name: '', email: '' };
    });
  }

  return {
    EXPIRED: EXPIRED,
    start: start,
    signIn: signIn,
    token: token,
    signOut: function () { clear(); },
    who: function () { return state ? { name: state.name, email: state.email } : null; }
  };
})();
