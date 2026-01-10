
#lang racket

;; Generate a PowerPoint .pptx with:
;;  - Slide 1: Title centered "My Presentation"
;;  - Slide 2: Title "First page" and 4 bullets that (aim to) appear progressively
;;  - Slide 3: Title "the end" and a centered PNG of a tabby cat
;;  - All slides: Persian carpet background image
;;
;; This builds a valid OOXML PresentationML package and zips it into .pptx.
;; References:
;;   - Office Open XML structure & parts: https://learn.microsoft.com/en-us/office/open-xml/presentation/structure-of-a-presentationml-document
;;   - Anatomy of a .pptx package: http://www.officeopenxml.com/anatomyofOOXML-pptx.php
;;   - Slide background via <p:bgPr>/<a:blipFill>: http://officeopenxml.com/prSlide-background.php
;;   - Pictures via <p:pic>/<a:blip r:embed=...>: https://ooxml.info/docs/l/l.3/l.3.5/l.4.10.2/
;;   - Animation/timing overview: https://learn.microsoft.com/en-us/office/open-xml/presentation/working-with-animation
;;   - Racket ZIP: https://docs.racket-lang.org/file/zip.html

(require racket/path
         racket/file
         racket/date
         racket/string
         file/zip)

;; ---- USER: set these to your local files ----
(define CARPET_PATH (build-path (current-directory) "carpet.jpg")) ; or .png
(define CAT_PATH     (build-path (current-directory) "tabby.png"))

(unless (file-exists? CARPET_PATH)
  (error 'setup "Persian carpet image not found at ~a" CARPET_PATH))
(unless (file-exists? CAT_PATH)
  (error 'setup "Tabby cat PNG not found at ~a" CAT_PATH))

;; Output file
(define OUT-PPTX (build-path (current-directory) "my-presentation.pptx"))

;; Work directory for building the package
(define WORK (build-path (current-directory) "pptx-build"))
(define (ensure-dir p) (unless (directory-exists? p) (make-directory* p)))

;; Folder layout (minimal but valid)
;; Root
(define ROOT WORK)
(define _rels     (build-path ROOT "_rels"))
(define docProps  (build-path ROOT "docProps"))
(define ppt       (build-path ROOT "ppt"))
(define ppt_rels  (build-path ppt "_rels"))
(define slides    (build-path ppt "slides"))
(define slides_rels (build-path slides "_rels"))
(define media     (build-path ppt "media"))
(define sldLayouts (build-path ppt "slideLayouts"))
(define sldLayouts_rels (build-path sldLayouts "_rels"))
(define sldMasters (build-path ppt "slideMasters"))
(define sldMasters_rels (build-path sldMasters "_rels"))
(define theme     (build-path ppt "theme"))

(for ([d (list ROOT _rels docProps ppt ppt_rels slides slides_rels
               media sldLayouts sldLayouts_rels sldMasters sldMasters_rels theme)])
  (ensure-dir d))

;; ---- Utilities ----
(define (write-bytes-file path bytes)
  (call-with-output-file path
    (lambda (out)
      (write-bytes bytes out))
    #:exists 'replace))

(define (write-xml path xml-str)
  (call-with-output-file path
    (lambda (out) (display xml-str out))
    #:exists 'replace))

(define (copy-into from to-name)
  (define dest (build-path media to-name))
  (copy-file from dest #:exists 'replace)
  to-name)

;; ---- Media ----
;; We'll embed the Persian carpet as "carpet.jpg" (or png) and the cat as "tabby.png"
(define carpet-ext (path-extension CARPET_PATH))
(define carpet-name (format "carpet.~a" (if carpet-ext (path-extension CARPET_PATH) "jpg")))
(define cat-name    "tabby.png")

(copy-into CARPET_PATH carpet-name)
(copy-into CAT_PATH    cat-name)

;; ---- [Content_Types].xml ----
;; Must list overrides for parts and default content types for image formats, etc.
(define content-types
  (string-append
   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
   "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">\n"
   "  <Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>\n"
   "  <Default Extension=\"xml\"  ContentType=\"application/xml\"/>\n"
   "  <Default Extension=\"png\"  ContentType=\"image/png\"/>\n"
   "  <Default Extension=\"jpg\"  ContentType=\"image/jpeg\"/>\n"
   "  <Default Extension=\"jpeg\" ContentType=\"image/jpeg\"/>\n"
   "  <Override PartName=\"/ppt/presentation.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml\"/>\n"
   "  <Override PartName=\"/ppt/slides/slide1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>\n"
   "  <Override PartName=\"/ppt/slides/slide2.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>\n"
   "  <Override PartName=\"/ppt/slides/slide3.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slide+xml\"/>\n"
   "  <Override PartName=\"/ppt/slideMasters/slideMaster1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml\"/>\n"
   "  <Override PartName=\"/ppt/slideLayouts/slideLayout1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml\"/>\n"
   "  <Override PartName=\"/ppt/theme/theme1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.theme+xml\"/>\n"
   "  <Override PartName=\"/docProps/core.xml\" ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>\n"
   "  <Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>\n"
   "</Types>\n"))

(write-xml (build-path ROOT "[Content_Types].xml") content-types)

;; ---- _rels/.rels ----
(define root-rels
  (string-append
   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
   "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
   "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"/ppt/presentation.xml\"/>\n"
   "  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" Target=\"/docProps/core.xml\"/>\n"
   "  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"/docProps/app.xml\"/>\n"
   "</Relationships>\n"))
(write-xml (build-path _rels ".rels") root-rels)

;; ---- docProps/core.xml & app.xml ----
(define now (current-seconds))
(define core-xml
  (string-append
   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
   "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" "
   "xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" "
   "xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n"
   "  <dc:title>My Presentation</dc:title>\n"
   "  <dc:creator>Racket program</dc:creator>\n"
   "  <cp:lastModifiedBy>Racket program</cp:lastModifiedBy>\n"
   "  <dcterms:created xsi:type=\"dcterms:W3CDTF\">" (date->string (seconds->date now) #t) "</dcterms:created>\n"
   "  <dcterms:modified xsi:type=\"dcterms:W3CDTF\">" (date->string (seconds->date now) #t) "</dcterms:modified>\n"
   "</cp:coreProperties>\n"))
(write-xml (build-path docProps "core.xml") core-xml)

(define app-xml
  "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n
   <Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"\n
               xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">\n
     <Application>Racket</Application>\n
     <DocSecurity>0</DocSecurity>\n
     <ScaleCrop>false</ScaleCrop>\n
     <Slides>3</Slides>\n
     <Notes>0</Notes>\n
   </Properties>\n")
(write-xml (build-path docProps "app.xml") app-xml)

;; ---- ppt/presentation.xml ----
(define presentation-xml
  (string-append
   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
   "<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"\n"
   "                xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"\n"
   "                xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">\n"
   "  <p:sldMasterIdLst>\n"
   "    <p:sldMasterId id=\"2147483648\" r:id=\"rId1\"/>\n"
   "  </p:sldMasterIdLst>\n"
   "  <p:sldIdLst>\n"
   "    <p:sldId id=\"256\" r:id=\"rId2\"/>\n"
   "    <p:sldId id=\"257\" r:id=\"rId3\"/>\n"
   "    <p:sldId id=\"258\" r:id=\"rId4\"/>\n"
   "  </p:sldIdLst>\n"
   "  <p:sldSz cx=\"9144000\" cy=\"6858000\" type=\"screen4x3\"/>\n"
   "  <p:notesSz cx=\"6858000\" cy=\"9144000\"/>\n"
   "</p:presentation>\n"))
(write-xml (build-path ppt "presentation.xml") presentation-xml)

(define presentation-rels
  (string-append
   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
   "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
   "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster\" Target=\"slideMasters/slideMaster1.xml\"/>\n"
   "  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide1.xml\"/>\n"
   "  <Relationship Id=\"rId3\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide2.xml\"/>\n"
   "  <Relationship Id=\"rId4\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide\" Target=\"slides/slide3.xml\"/>\n"
   "  <Relationship Id=\"rId5\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme\" Target=\"theme/theme1.xml\"/>\n"
   "</Relationships>\n"))
(write-xml (build-path ppt_rels "presentation.xml.rels") presentation-rels)

;; ---- theme (minimal placeholder) ----
(define theme-xml
  (string-append
   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
   "<a:theme xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"MinimalTheme\">\n"
   "  <a:themeElements>\n"
   "    <a:clrScheme name=\"Office\">\n"
   "      <a:dk1><a:srgbClr val=\"000000\"/></a:dk1>\n"
   "      <a:lt1><a:srgbClr val=\"FFFFFF\"/></a:lt1>\n"
   "      <a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2>\n"
   "      <a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2>\n"
   "      <a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1>\n"
   "      <a:accent2><a:srgbClr val=\"C0504D\"/></a:accent2>\n"
   "      <a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3>\n"
   "      <a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4>\n"
   "      <a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5>\n"
   "      <a:accent6><a:srgbClr val=\"F79646\"/></a:accent6>\n"
   "      <a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink>\n"
   "      <a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink>\n"
   "    </a:clrScheme>\n"
   "    <a:fontScheme name=\"Minor\"><a:majorFont/><a:minorFont/></a:fontScheme>\n"
   "    <a:fmtScheme name=\"Simple\"><a:fillStyleLst/><a:lnStyleLst/><a:effectStyleLst/><a:bgFillStyleLst/></a:fmtScheme>\n"
   "  </a:themeElements>\n"
   "</a:theme>\n"))
(write-xml (build-path theme "theme1.xml") theme-xml)

;; ---- Slide Master & Layout (very small) ----
(define slide-master-xml
  (string-append
   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
   "<p:sldMaster xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
   "  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm/></p:grpSpPr></p:spTree></p:cSld>\n"
   "  <p:clrMap bg1=\"dk1\" tx1=\"lt1\" bg2=\"dk2\" tx2=\"lt2\" accent1=\"accent1\" accent2=\"accent2\" accent3=\"accent3\" accent4=\"accent4\" accent5=\"accent5\" accent6=\"accent6\" hlink=\"hlink\" folHlink=\"folHlink\"/>\n"
   "  <p:sldLayoutIdLst><p:sldLayoutId id=\"1\" r:id=\"rId1\"/></p:sldLayoutIdLst>\n"
   "</p:sldMaster>\n"))
(write-xml (build-path sldMasters "slideMaster1.xml") slide-master-xml)

(define slide-master-rels
  (string-append
   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
   "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
   "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout\" Target=\"../slideLayouts/slideLayout1.xml\"/>\n"
   "</Relationships>\n"))
(write-xml (build-path sldMasters_rels "slideMaster1.xml.rels") slide-master-rels)

(define slide-layout-xml
  (string-append
   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
   "<p:sldLayout xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" type=\"title\" preserve=\"1\">\n"
   "  <p:cSld name=\"Title Layout\"><p:spTree><p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm/></p:grpSpPr></p:spTree></p:cSld>\n"
   "  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>\n"
   "</p:sldLayout>\n"))
(write-xml (build-path sldLayouts "slideLayout1.xml") slide-layout-xml)

(write-xml (build-path sldLayouts_rels "slideLayout1.xml.rels")
           "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"/>")

;; ---- Helper snippets ----

;; A: centered title shape at slide center
(define (title-shape text spid)
  (format
   (string-append
    "<p:sp>\n"
    "  <p:nvSpPr>\n"
    "    <p:cNvPr id=\"~a\" name=\"Title\"/>\n"
    "    <p:cNvSpPr/>\n"
    "    <p:nvPr/>\n"
    "  </p:nvSpPr>\n"
    "  <p:spPr>\n"
    "    <a:xfrm>\n"
    "      <a:off x=\"914400\" y=\"2286000\"/>\n"      ; position
    "      <a:ext cx=\"7315200\" cy=\"1143000\"/>\n"  ; size
    "    </a:xfrm>\n"
    "    <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>\n"
    "    <a:noFill/>\n"
    "    <a:ln><a:noFill/></a:ln>\n"
    "  </p:spPr>\n"
    "  <p:txBody>\n"
    "    <a:bodyPr/>\n"
    "    <a:lstStyle/>\n"
    "    <a:p>\n"
    "      <a:pPr algn=\"ctr\"/>\n"
    "      <a:r><a:rPr lang=\"en-GB\" sz=\"4400\" b=\"1\"/><a:t>~a</a:t></a:r>\n"
    "      <a:endParaRPr/>\n"
    "    </a:p>\n"
    "  </p:txBody>\n"
    "</p:sp>\n")
   spid text))

;; B: title at top
(define (top-title-shape text spid)
  (format
   (string-append
    "<p:sp>\n"
    "  <p:nvSpPr>\n"
    "    <p:cNvPr id=\"~a\" name=\"SlideTitle\"/>\n"
    "    <p:cNvSpPr/>\n"
    "    <p:nvPr/>\n"
    "  </p:nvSpPr>\n"
    "  <p:spPr>\n"
    "    <a:xfrm><a:off x=\"457200\" y=\"457200\"/><a:ext cx=\"8229600\" cy=\"914400\"/></a:xfrm>\n"
    "    <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>\n"
    "    <a:noFill/><a:ln><a:noFill/></a:ln>\n"
    "  </p:spPr>\n"
    "  <p:txBody>\n"
    "    <a:bodyPr/>\n"
    "    <a:lstStyle/>\n"
    "    <a:p><a:pPr algn=\"ctr\"/><a:r><a:rPr lang=\"en-GB\" sz=\"3600\" b=\"1\"/><a:t>~a</a:t></a:r><a:endParaRPr/></a:p>\n"
    "  </p:txBody>\n"
    "</p:sp>\n")
   spid text))

;; C: bullet list body (four bullets)
(define (bullet-body-shape bullets spid)
  (define paras
    (string-join
     (for/list ([b bullets] [i (in-naturals 0)])
       (format
        (string-append
         "<a:p>\n"
         "  <a:pPr lvl=\"0\">\n"
         "    <a:buChar char=\"•\"/>\n"
         "    <a:buFont typeface=\"Arial\"/>\n"
         "  </a:pPr>\n"
         "  <a:r><a:rPr lang=\"en-GB\" sz=\"2800\"/><a:t>~a</a:t></a:r>\n"
         "  <a:endParaRPr/>\n"
         "</a:p>\n")
        b))
     ""))

  (format
   (string-append
    "<p:sp>\n"
    "  <p:nvSpPr>\n"
    "    <p:cNvPr id=\"~a\" name=\"BulletList\"/>\n"
    "    <p:cNvSpPr/>\n"
    "    <p:nvPr/>\n"
    "  </p:nvSpPr>\n"
    "  <p:spPr>\n"
    "    <a:xfrm><a:off x=\"685800\" y=\"1828800\"/><a:ext cx=\"7772400\" cy=\"3657600\"/></a:xfrm>\n"
    "    <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>\n"
    "    <a:noFill/><a:ln><a:noFill/></a:ln>\n"
    "  </p:spPr>\n"
    "  <p:txBody>\n"
    "    <a:bodyPr/>\n"
    "    <a:lstStyle/>\n"
    "    ~a\n"
    "  </p:txBody>\n"
    "</p:sp>\n")
   spid paras))

;; D: slide background as Persian carpet image (picture fill)
;;    Uses <p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdX"/>
(define (slide-bg rId)
  (format
   (string-append
    "<p:bg>\n"
    "  <p:bgPr>\n"
    "    <a:blipFill>\n"
    "      <a:blip r:embed=\"~a\"/>\n"
    "      <a:stretch><a:fillRect/></a:stretch>\n"
    "    </a:blipFill>\n"
    "    <a:effectLst/>\n"
    "  </p:bgPr>\n"
    "</p:bg>\n")
   rId))

;; E: centered picture (cat)
(define (picture-shape rId spid)
  (format
   (string-append
    "<p:pic>\n"
    "  <p:nvPicPr>\n"
    "    <p:cNvPr id=\"~a\" name=\"Cat\"/>\n"
    "    <p:cNvPicPr><a:picLocks noChangeAspect=\"1\"/></p:cNvPicPr>\n"
    "    <p:nvPr/>\n"
    "  </p:nvPicPr>\n"
    "  <p:blipFill>\n"
    "    <a:blip r:embed=\"~a\"/>\n"
    "    <a:stretch><a:fillRect/></a:stretch>\n"
    "  </p:blipFill>\n"
    "  <p:spPr>\n"
    "    <a:xfrm>\n"
    "      <a:off x=\"2743200\" y=\"1828800\"/>\n"      ; roughly centered
    "      <a:ext cx=\"3657600\" cy=\"2743200\"/>\n"
    "    </a:xfrm>\n"
    "    <a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>\n"
    "    <a:noFill/><a:ln><a:noFill/></a:ln>\n"
    "  </p:spPr>\n"
    "</p:pic>\n")
   spid rId))

;; ---- Slides ----

;; Slide 1: title centered + Persian carpet background
(define slide1-xml
  (string-append
   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
   "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
   "  <p:cSld>\n"
   (slide-bg "rId1")
   "    <p:spTree>\n"
   "      <p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>\n"
   "      <p:grpSpPr><a:xfrm/></p:grpSpPr>\n"
   (title-shape "My Presentation" 2)
   "    </p:spTree>\n"
   "  </p:cSld>\n"
   "  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>\n"
   "</p:sld>\n"))
(write-xml (build-path slides "slide1.xml") slide1-xml)

(write-xml (build-path slides_rels "slide1.xml.rels")
           (string-append
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
            "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/" carpet-name "\"/>\n"
            "</Relationships>\n"))

;; Slide 2: title top + bullet list + carpet background
(define bullets
  (list "My first slide"
        "my wonderful coding assistant"
        "my cat"
        "do not let the dog in the house"))

(define slide2-shapes
  (string-append
   (top-title-shape "First page" 2)
   (bullet-body-shape bullets 3)))

;; Minimal timing: sequence four on-click steps targeting the BulletList text by paragraph.
;; NOTE: Animation markup is complex; this uses a general <p:timing> with a mainSeq of four
;; child nodes that many PowerPoint versions map to “By Paragraph” appearance.
(define slide2-timing
  (string-append
   "  <p:timing>\n"
   "    <p:tnLst>\n"
   "      <p:par>\n"
   "        <p:cTn id=\"1\" dur=\"indefinite\" nodeType=\"tmRoot\"/>\n"
   "        <p:childTnLst>\n"
   "          <p:seq concurrent=\"0\" nextAc=\"seek\">\n"
   "            <p:cTn id=\"2\" dur=\"indefinite\" nodeType=\"mainSeq\"/>\n"
   "            <p:childTnLst>\n"
   ;; 4 “appear” steps tied to the BulletList (spid=3), one per paragraph (1..4)
   (string-join
    (for/list ([i (in-range 1 5)])
      (format
       (string-append
        "              <p:par>\n"
        "                <p:cTn id=\"~a\" dur=\"500\" fill=\"hold\">\n"
        "                  <p:stCondLst><p:cond evt=\"onBegin\"/></p:stCondLst>\n"
        "                  <p:tgtEl>\n"
        "                    <p:spTgt spid=\"3\">\n"
        "                      <p:txEl><p:pRg st=\"~a\" end=\"~a\"/></p:txEl>\n"
        "                    </p:spTgt>\n"
        "                  </p:tgtEl>\n"
        "                  <p:anim valueType=\"num\">\n"
        "                    <p:cBhvr>\n"
        "                      <p:cTn id=\"~a01\" dur=\"500\"/>\n"
        "                      <p:attrNameLst>\n"
        "                        <p:attrName>style.visibility</p:attrName>\n"
        "                      </p:attrNameLst>\n"
        "                    </p:cBhvr>\n"
        "                  </p:anim>\n"
        "                </p:cTn>\n"
        "              </p:par>\n")
       (+ 10 i) i i))
    "")
   "            </p:childTnLst>\n"
   "          </p:seq>\n"
   "        </p:childTnLst>\n"
   "      </p:par>\n"
   "    </p:tnLst>\n"
   "  </p:timing>\n"))

(define slide2-xml
  (string-append
   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
   "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
   "  <p:cSld>\n"
   (slide-bg "rId1")
   "    <p:spTree>\n"
   "      <p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>\n"
   "      <p:grpSpPr><a:xfrm/></p:grpSpPr>\n"
   slide2-shapes
   "    </p:spTree>\n"
   "  </p:cSld>\n"
   "  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>\n"
   slide2-timing
   "</p:sld>\n"))
(write-xml (build-path slides "slide2.xml") slide2-xml)

(write-xml (build-path slides_rels "slide2.xml.rels")
           (string-append
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
            "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/" carpet-name "\"/>\n"
            "</Relationships>\n"))

;; Slide 3: title top + centered PNG cat + carpet background
(define slide3-xml
  (string-append
   "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
   "<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\n"
   "  <p:cSld>\n"
   (slide-bg "rId1")  ;; background image
   "    <p:spTree>\n"
   "      <p:nvGrpSpPr><p:cNvPr id=\"1\" name=\"\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>\n"
   "      <p:grpSpPr><a:xfrm/></p:grpSpPr>\n"
   (top-title-shape "the end" 2)
   (picture-shape "rId2" 3)
   "    </p:spTree>\n"
   "  </p:cSld>\n"
   "  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>\n"
   "</p:sld>\n"))
(write-xml (build-path slides "slide3.xml") slide3-xml)

(write-xml (build-path slides_rels "slide3.xml.rels")
           (string-append
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n"
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n"
            "  <Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/" carpet-name "\"/>\n"
            "  <Relationship Id=\"rId2\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/" cat-name "\"/>\n"
            "</Relationships>\n"))

;; ---- Zip into .pptx ----
;; Racket's `zip` packs files into a ZIP; a .pptx is just that ZIP with this structure.
;; (It zips paths relative to current directory.)
(define (collect-relative-paths base)
  ;; walk directory, return list of relative paths
  (for/list ([p (in-directory base)])
    (define rel (simplify-path (path-replace-prefix p base "")))
    rel))

;; Move into WORK so the archive contains the correct paths (no absolute prefixes)
(parameterize ([current-directory WORK])
  (zip OUT-PPTX (collect-relative-paths ".")))

(printf "Wrote ~a\n" OUT-PPTX)
