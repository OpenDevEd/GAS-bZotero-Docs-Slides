function addCurrentKeys(urlpart) {
  // New 2021-04-24
  let newUrl = redirectTarget + urlpart;

  let currentZoteroItemKey = getDocumentPropertyString('zotero_item');
  if (currentZoteroItemKey != null) {
    const zoteroItemKeyParts = currentZoteroItemKey.split('/');
    const zoteroItemKeyParameters = zoteroItemKeyParts[4] + ':' + zoteroItemKeyParts[6];
    newUrl = replaceAddParameter(newUrl, 'src', zoteroItemKeyParameters);
  }

  const currentZoteroCollectionKey = getDocumentPropertyString('zotero_collection_key');
  if (currentZoteroCollectionKey != null) {
    const zoteroCollectionKeyParts = currentZoteroCollectionKey.split('/');
    zoteroCollectionKey = zoteroCollectionKeyParts[6];
    newUrl = replaceAddParameter(newUrl, 'collection', zoteroCollectionKey);
  }

  newUrl = replaceAddParameter(newUrl, 'openin', 'zoteroapp');
  // End. New 2021-04-24
  return newUrl;
}
