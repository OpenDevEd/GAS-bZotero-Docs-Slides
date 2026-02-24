function onInstall(e) {
  onOpen(e);
  addInstallationRecord();
}

function minifyCitations() {
  if (HOST_APP == 'docs') {
    minifyCitationsDocs();
  } else {
    packZoteroSlides('minifyCitations');
  }
  addUsageTrackingRecord('minifyCitations');
};

function maxifyCitations() {
  if (HOST_APP == 'docs') {
    maxifyCitationsDocs();
  } else {
    packZoteroSlides('maxifyCitations');
  }
  addUsageTrackingRecord('maxifyCitations');
};

function unfyCitations() {
  if (HOST_APP == 'docs') {
    unfyCitationsDocs();
  } else {
    packZoteroSlides('unfyCitations');
  }
  addUsageTrackingRecord('unfyCitations');
};

function unpackCombined() {
  if (HOST_APP == 'docs') {
    restoreZoteroLinks();
    zoteroUnpackCall(false);
  } else {
    unpackZoteroSlides(false);
  }
  addUsageTrackingRecord('unpackCombined');
};

function unpackCombinedWarning() {
  if (HOST_APP == 'docs') {
    restoreZoteroLinks();
    zoteroUnpackCall(true);
  } else {
    unpackZoteroSlides(true);
  }
  addUsageTrackingRecord('unpackCombinedWarning');
};

function packZoteroCall() {
  if (HOST_APP == 'docs') {
    zoteroPackUnpack(true, false, null);
  } else {
    packZoteroSlides('packZotero');
  }
  addUsageTrackingRecord('packZoteroCall');
};

function zoteroUnpackCall(warning) {
  zoteroPackUnpack(false, false, warning);
};

function packZoteroSelectiveCall() {
  if (HOST_APP == 'docs') {
    zoteroPackUnpack(true, true, null);
  } else {
    packZoteroSelectiveCallSlides();
  }
  addUsageTrackingRecord('packZoteroSelectiveCall');
};