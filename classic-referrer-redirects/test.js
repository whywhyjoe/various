// node test.js — stdlib-only tests for the pure matching logic.
'use strict';
var assert = require('assert');
var cr = require('./classic-redirect.js');

var config = {
  oldSiteFragment: '/sites/TandO-Classic/FCU',
  newSiteBase: '/sites/TandO-FCU',
  redirects: {
    '': 'SitePages/Home.aspx',
    'NWWSH': 'SitePages/NWWSH.aspx',
    'NWWSH/blah.aspx': 'SitePages/NWWSH-Blah.aspx',
    'Training': 'https://contoso.sharepoint.com/sites/Other/Training.aspx'
  }
};
var OLD = 'https://contoso.sharepoint.com/sites/TandO-Classic/FCU';
var HERE = '/sites/TandO-FCU/SitePages/Landing.aspx';

function resolves(referrer, expected, cfg) {
  assert.strictEqual(cr.resolveRedirect(cfg || config, referrer, HERE), expected,
    'referrer: ' + referrer);
}

// No referrer / referrer outside the old site -> stay put.
resolves('', null);
resolves('https://contoso.sharepoint.com/sites/SomethingElse/page.aspx', null);
// Fragment must match a whole path segment, not a prefix.
resolves('https://contoso.sharepoint.com/sites/TandO-Classic/FCU-other/x.aspx', null);

// Section (no page name) vs specific page, per the NWWSH example.
resolves(OLD + '/NWWSH',
  '/sites/TandO-FCU/SitePages/NWWSH.aspx?classicRedirect=1&classicRedirectFrom=' +
  encodeURIComponent('/sites/TandO-Classic/FCU/NWWSH'));
resolves(OLD + '/NWWSH/blah.aspx',
  '/sites/TandO-FCU/SitePages/NWWSH-Blah.aspx?classicRedirect=1&classicRedirectFrom=' +
  encodeURIComponent('/sites/TandO-Classic/FCU/NWWSH/blah.aspx'));

// Trailing slash, query strings, and case differences don't matter.
resolves(OLD + '/NWWSH/',
  '/sites/TandO-FCU/SitePages/NWWSH.aspx?classicRedirect=1&classicRedirectFrom=' +
  encodeURIComponent('/sites/TandO-Classic/FCU/NWWSH/'));
resolves(OLD.toLowerCase() + '/nwwsh/BLAH.ASPX?tab=2#frag',
  '/sites/TandO-FCU/SitePages/NWWSH-Blah.aspx?classicRedirect=1&classicRedirectFrom=' +
  encodeURIComponent('/sites/tando-classic/fcu/nwwsh/BLAH.ASPX'));

// The classic welcome page counts as its folder.
resolves(OLD + '/NWWSH/default.aspx',
  '/sites/TandO-FCU/SitePages/NWWSH.aspx?classicRedirect=1&classicRedirectFrom=' +
  encodeURIComponent('/sites/TandO-Classic/FCU/NWWSH/default.aspx'));

// The old site root itself maps via the "" key.
resolves(OLD,
  '/sites/TandO-FCU/SitePages/Home.aspx?classicRedirect=1&classicRedirectFrom=' +
  encodeURIComponent('/sites/TandO-Classic/FCU'));

// Absolute destinations pass through untouched (aside from the signal).
resolves(OLD + '/Training',
  'https://contoso.sharepoint.com/sites/Other/Training.aspx?classicRedirect=1&classicRedirectFrom=' +
  encodeURIComponent('/sites/TandO-Classic/FCU/Training'));

// Unmapped old page: null without a fallback, fallback when configured.
resolves(OLD + '/no-such-page.aspx', null);
var withFallback = JSON.parse(JSON.stringify(config));
withFallback.fallback = 'SitePages/Migration-Info.aspx?src=old';
resolves(OLD + '/no-such-page.aspx',
  '/sites/TandO-FCU/SitePages/Migration-Info.aspx?src=old&classicRedirect=1&classicRedirectFrom=' +
  encodeURIComponent('/sites/TandO-Classic/FCU/no-such-page.aspx'),
  withFallback);

// A custom signalParam is used for both query params.
var renamed = JSON.parse(JSON.stringify(config));
renamed.signalParam = 'wasClassic';
resolves(OLD + '/NWWSH',
  '/sites/TandO-FCU/SitePages/NWWSH.aspx?wasClassic=1&wasClassicFrom=' +
  encodeURIComponent('/sites/TandO-Classic/FCU/NWWSH'),
  renamed);

// Never redirect the landing page to itself.
var selfMap = JSON.parse(JSON.stringify(config));
selfMap.redirects['NWWSH'] = 'SitePages/Landing.aspx';
resolves(OLD + '/NWWSH', null, selfMap);

console.log('classic-referrer-redirects: all tests passed');
