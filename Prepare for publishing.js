function prepareForPublishing() {

  let targetRefLinks = getDocumentPropertyString('target_ref_links');
 
  if (targetRefLinks != 'kerko') {
    setDocumentPropertyString('target_ref_links', 'kerko');
    onOpen();
  }

  if (HOST_APP == 'docs') {
    universalInsertUpdateBibliography(true, true, true);
  } else {
    insertUpdateBibliographySlides(true, true, true);
  }

  removeUnderlineFromHyperlinks();
}