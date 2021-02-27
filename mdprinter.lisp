;;;; mdprinter.lisp

(cl:in-package #:3bmd2docx)

;;;; 3bmd markdown printer

#|

Use:

(3bmd:parse-string-and-print-to-stream string stream :format :wml)

(3bmd:parse-and-print-to-stream file stream :format :wml)

(3bmd:print-doc-to-stream doc stream :format :wml)

Note: relies on 3bmd::expand-tabs to get rid of tabs

|#

(defparameter *document* nil)
(defparameter *numbering-definitions* nil)
(defparameter *style-definitions* nil)
(defparameter *md-infile* nil)

(defparameter *html-is-wml* nil)

(defparameter *in-code* nil)
(defparameter *in-block-quote* nil)
(defparameter *in-paragraph* nil)
(defparameter *para-justification* nil) ;; stack of '3bmd-grammar::(left center right)
(defparameter *in-run* nil)
(defparameter *in-text* nil)
(defparameter *run-props* '()) ;; stack of :emph, :strong, :code and :hyperlink
(defparameter *in-list* nil) ;; stack of :counted-list and :bullet-list
(defparameter *new-list* nil) ;; t when starting a new (counted) list
(defparameter *very-first-list* t) ;; t when it's the first list in the doc

(defun preservep (string)
  (or (alexandria:starts-with #\Space string :test #'char=)
      (alexandria:ends-with #\Space string :test #'char=)))

(defun sym-jc (sym)
  (ecase sym
    ((3bmd-grammar::left) "start")
    ((3bmd-grammar::center) "center")
    ((3bmd-grammar::right) "end")))

(defun open-paragraph (stream &key style-name ilvl numid)
  (when *in-paragraph* (close-paragraph stream))
  (setf *in-paragraph* t)
  (write-string "<w:p>" stream)
  (when (or style-name *para-justification*)
    (write-string "<w:pPr>" stream)
    (when style-name
      (format stream "<w:pStyle w:val=\"~A\" />" style-name))
    (when (and ilvl numid)
      (format stream "<w:numPr><w:ilvl w:val=\"~D\" /><w:numId w:val=\"~D\" /></w:numPr>"
	      ilvl numid))
    (when *para-justification*
      (format stream "<w:jc w:val=\"~A\" />" (sym-jc (first *para-justification*))))
    (write-string "</w:pPr>" stream)))

(defun close-paragraph (stream)
  (when *in-run* (close-run stream))
  (write-string "</w:p>" stream)
  (setf *in-paragraph* nil)
  (terpri stream)) ;;FIXME: readability

(defun open-run (stream)
  (write-string "<w:r>" stream)
  (when *run-props*
    (write-string "<w:rPr>" stream)
    (dolist (rp *run-props*)
      (ecase rp
	(:emph (write-string "<w:i w:val=\"true\" /><w:iCs w:val=\"true\" />" stream))
	(:strong (write-string "<w:b w:val=\"true\" /><w:bCs w:val=\"true\" />" stream))
	(:code (write-string "<w:rStyle w:val=\"mdcode\" />" stream));"<w:rFonts w:ascii=\"Consolas\" />" stream))
	(:hyperlink (write-string "<w:rStyle w:val=\"mdlink\" />" stream)))) 
    (write-string "</w:rPr>" stream))
  (setf *in-run* t))

(defun close-run (stream)
  (when *in-text*
    (close-text stream))
  (write-string "</w:r>" stream)
  (setf *in-run* nil))

(defun open-text (stream string)
  (unless *in-run*
    (open-run stream))
  (when *in-text*
    (close-text stream))
  (format stream "<w:t~@[ xml:space=\"preserve\"~]>" (preservep string))
  (setf *in-text* t))

(defun close-text (stream)
  (format stream "</w:t>")
  (setf *in-text* nil))

(defparameter *references* nil)

(defun print-escaped (string stream)
  (loop for c across string
     when (eql c #\&) do (write-string "&amp;" stream)
     else when (eql c #\<) do (write-string "&lt;" stream)
     else when (eql c #\>) do (write-string "&gt;" stream)
     else do (write-char c stream)))

(defun escape-string (string)
  (with-output-to-string (s)
    (print-escaped string s)))

(defgeneric print-tagged-element (tag stream rest))

(defmethod print-tagged-element ((tag (eql :ellipsis)) stream rest)
  (if *in-code*
      (dolist (elem rest) (print-element elem stream))
      (print-element (string #\horizontal_ellipsis) stream)))

(defmethod print-tagged-element ((tag (eql :single-quoted)) stream rest)
  (if *in-code*
      (print-element "'" stream)
      (print-element (string #\left_single_quotation_mark) stream))
  (dolist (elem rest) (print-element elem stream))
  (if *in-code*
      (print-element "'" stream)
      (print-element (string #\right_single_quotation_mark) stream)))

(defmethod print-tagged-element ((tag (eql :double-quoted)) stream rest)
  (if *in-code*
      (print-element "\"" stream)
      (print-element (string #\left_double_quotation_mark) stream))
  (dolist (elem rest) (print-element elem stream))
  (if *in-code*
      (print-element "\"" stream)
      (print-element (string #\right_double_quotation_mark) stream)))

(defmacro define-smart-quote-entity (name replacement)
  `(defmethod print-tagged-element ((tag (eql ,name)) stream rest)
     (if *in-code*
	 (dolist (elem rest) (print-element elem stream))
	 (print-element (string ,replacement) stream))))

(define-smart-quote-entity :em-dash #\em_dash)
(define-smart-quote-entity :en-dash #\en_dash)
(define-smart-quote-entity :left-right-single-arrow #\left_right_arrow)
(define-smart-quote-entity :left-single-arrow #\leftwards_arrow)
(define-smart-quote-entity :right-single-arrow #\rightwards_arrow)
(define-smart-quote-entity :left-right-double-arrow #\left_right_double_arrow)
(define-smart-quote-entity :left-double-arrow #\leftwards_double_arrow)
(define-smart-quote-entity :right-double-arrow #\rightwards_double_arrow)

(defmethod print-tagged-element ((tag (eql :line-break)) stream rest)
  (when *in-text* (close-text stream))
  (unless *in-run* (open-run stream))
  (write-string "<w:cr />" stream))

(defmethod print-tagged-element ((tag (eql :horizontal-rule)) stream rest)
  (write-string "<w:p><w:pPr><w:pBdr><w:bottom w:val=\"single\" w:sz=\"12\" w:color=\"auto\" w:space=\"1\" /></w:pBdr></w:pPr></w:p>" stream)
  (terpri stream))

(defmethod print-tagged-element ((tag (eql :paragraph)) stream rest)
  (let ((paragraph-style (cond (*in-block-quote* "mdquote")
			      ; (*in-code* "mdcode")
			       (t nil))))
    (unless *in-list*
      (open-paragraph stream :style-name paragraph-style)))
  (dolist (elem rest) (print-element elem stream))
  (close-paragraph stream))

(defmethod print-tagged-element ((tag (eql :block-quote)) stream rest)
  (let ((*in-block-quote* t))
    (dolist (elem rest) (print-element elem stream))))

(defmethod print-tagged-element ((tag (eql :heading)) stream rest)
  (open-paragraph stream :style-name (format nil "mdheading~D" (getf rest :level)))
  (dolist (elem (getf rest :contents)) (print-element elem stream))
  (close-paragraph stream))

(defmethod print-tagged-element ((tag (eql :counted-list)) stream rest)
    (if *very-first-list*
	(setf *very-first-list* nil)
	(setf *new-list* t))
  (push :counted-list *in-list*)
  (dolist (elem rest) (print-element elem stream))
  (pop *in-list*))

(defmethod print-tagged-element ((tag (eql :bullet-list)) stream rest)
  (push :bullet-list *in-list*)
  (dolist (elem rest) (print-element elem stream))
  (pop *in-list*))

(defun maybe-add-start-override ()
  (if *document*
      (let* ((nso (docxplora:make-numbering-start-override *numbering-definitions* 1))
	     (num-id (plump:attribute nso "w:numId"))
	     (numbering (docxplora:get-first-element-by-tag-name
			 (opc:xml-root *numbering-definitions*)
			 "w:numbering")))
	(plump:append-child numbering nso)
	num-id)
      2))

(defmethod print-tagged-element ((tag (eql :list-item)) stream rest)
  (let ((numid (ecase (first *in-list*)
		 (:counted-list (cond (*new-list*
				       (setf *new-list* nil)
				       (maybe-add-start-override))
				      (t 1)))
		 (:bullet-list 3)))
	(ilvl (1- (length *in-list*))))
    (open-paragraph stream :style-name "mdlistparagraph":ilvl ilvl :numid numid)
    (dolist (elem rest) (print-element elem stream))))

(defmethod print-tagged-element ((tag (eql :code)) stream rest)
  (when *in-run* (close-run stream))
  (push :code *run-props*)
  (push t *in-code*)
  (dolist (elem rest) (print-element elem stream))
  (when *in-run* (close-run stream))
  (pop *in-code*)
  (pop *run-props*))

(defmethod print-tagged-element ((tag (eql :emph)) stream rest)
  (when *in-run* (close-run stream))
  (push :emph *run-props*)
  (dolist (elem rest) (print-element elem stream))
  (when *in-run* (close-run stream))
  (pop *run-props*))

(defmethod print-tagged-element ((tag (eql :strong)) stream rest)
  (when *in-run* (close-run stream))
  (push :strong *run-props*)
  (dolist (elem rest) (print-element elem stream))
  (when *in-run* (close-run stream))
  (pop *run-props*))

(defun internalp (link)
  (alexandria:starts-with #\# link :test #'char=))

(defmethod print-tagged-element ((tag (eql :link)) stream rest)
  (when *in-run* (close-run stream))
  (let ((id (if *document*
		(opc:relationship-id
		 (docxplora:ensure-hyperlink *document* (car rest)))
		(car rest)))) ; Fake ID if not in document
    (format stream "<w:hyperlink w:id=\"~A\">" id))
  (push :hyperlink *run-props*)
  (print-element (car rest) stream)
  (when *in-run* (close-run stream))
  (pop *run-props*)
  (write-string "</w:hyperlink>" stream))

(defmethod print-tagged-element ((tag (eql :mailto)) stream rest)
  (when *in-run* (close-run stream))
  (let ((id (if *document*
		(opc:relationship-id
		 (docxplora:ensure-hyperlink *document*
					     (car rest)))
		(car rest))))
    (format stream "<w:hyperlink w:id=\"~A\">" id))
  (push :hyperlink *run-props*)
  (print-element (car rest) stream)
  (when *in-run* (close-run stream))
  (pop *run-props*)
  (write-string "</w:hyperlink>" stream))

(defmethod print-tagged-element ((tag (eql :explicit-link)) stream rest)
  (when *in-run* (close-run stream))
  (let ((source (getf rest :source)))
    (setf source (if (internalp source)
		     (format nil "w:anchor=\"~A\"" (subseq source 1))
		     (format nil "w:id=\"~A\""
			     (opc:relationship-id
			      (docxplora:ensure-hyperlink *document* source)))))
    (format stream "<w:hyperlink ~A ~@[w:tooltip=\"~A\"~]>"
	    source
	    (alexandria:when-let (title (getf rest :title))
	      (escape-string title))))
  (push :hyperlink *run-props*)
  (dolist (elem (getf rest :label)) (print-element elem stream))
  (when *in-run* (close-run stream))
  (pop *run-props*)
  (write-string "</w:hyperlink>" stream))

(defmethod print-tagged-element ((tag (eql :reference-link)) stream rest)
  (let* ((label (getf rest :label))
	 (def (or (getf rest :definition) label))
	 (ref (3bmd::lookup-reference def)))
    (cond (ref
	   (when *in-run* (close-run stream))
	   (let ((source (first ref)))
	     (setf source (if (internalp source)
			      (format nil "w:anchor=\"~A\"" (subseq source 1))
			      (format nil "w:id=\"~A\""
				      (opc:relationship-id
				       (docxplora:ensure-hyperlink *document* source)))))
	     (format stream "<w:hyperlink ~A ~@[w:tooltip=\"~A\"~]>"
		     source
		     (second ref)))
	   (push :hyperlink *run-props*)
	   (dolist (elem (getf rest :label)) (print-element elem stream))
	   (when *in-run* (close-run stream))
	   (pop *run-props*)
	   (write-string "</w:hyperlink>" stream))
	  (t
	   (print-element "[" stream)
	   (dolist (elem label) (print-element elem stream))
	   (print-element "]" stream)
	   (alexandria:when-let (tail (getf rest :tail))
	     (print-element tail stream))))))

(defmethod print-tagged-element ((tag (eql :image)) stream rest)
  (let* ((rest (cdr (first rest)))
	 (source (getf rest :source))
	 (label (getf rest :label))
	 #+(or)(title (getf rest :title)))
    (flet ((print-as-text ()
	     (print-element "![" stream)
	     (dolist (elem label) (print-element elem stream))
	     (print-element "]" stream)))
      (cond ((or (null *document*)
		 (null source)
		 (null
		  (ignore-errors
		    (imagesniff:image-type
		     (merge-pathnames source *md-infile*)))))
	     (print-as-text))
	    (t
	     (let ((im (docxplora:make-inline-image
			*document*
			(merge-pathnames source *md-infile*))))
	       (when *in-text* (close-text stream))
	       (plump:serialize im stream)))))))

(defun print-stuff (stuff stream)
  (when *in-run* (close-run stream))
  (dolist (elem stuff) (print-element elem stream))
  (when *in-run* (close-run stream)))

(defmethod print-tagged-element ((tag (eql :html)) stream rest)
  (when *html-is-wml*
    (format stream "~{~a~}" rest)))

(defmethod print-tagged-element ((tag (eql :raw-html)) stream rest)
  (when *html-is-wml*
    (format stream "~{~a~}" rest)))

(defmethod print-tagged-element ((tag (eql :entity)) stream rest) ;;FIXME check
  (dolist (string rest)
    (if (member string '("&lt;" "&amp;" "&gt;") :test #'string=)
	(write-string string stream)
	(write-string (plump:decode-entities string) stream))))

(defun print-verbatim (stream stuff)
  (dolist (string stuff)
    (let ((lines (split-sequence:split-sequence #\Newline string)))
      (dolist (line lines)
	(open-paragraph stream)
	(push :code *run-props*)
	(open-text stream line)
	(print-escaped line stream)
	(pop *run-props*)
	(close-run stream)
	(close-paragraph stream)))))

(defmethod print-tagged-element ((tag (eql :verbatim)) stream rest)
  (print-verbatim stream rest))

(defmethod print-tagged-element ((tag (eql :plain)) stream rest)
  (print-stuff rest stream))


(defmethod print-tagged-element ((tag (eql :reference)) stream rest)
  )

;;; tables

(defmethod print-tagged-element ((tag (eql '3bmd::table)) stream rest)
  (write-string "<w:tbl><w:tblPr><w:tblStyle w:val=\"mdtable\" /></w:tblPr>" stream)
  (dolist (row (getf rest :head)) (print-table-row stream row))
  (dolist (row (getf rest :body)) (print-table-row stream row))
  (write-string "</w:tbl><w:p/>" stream)) ; Add paragraph to avoid adjacent tables being joined

(defun print-table-row (stream row)
  (write-string "<w:tr>" stream)
  (dolist (cell row)
    (write-string "<w:tc>" stream)
    (when (third cell) (push (third cell) *para-justification*))
    (print-tagged-element :paragraph stream (rest (second cell)))
    (when (third cell) (pop *para-justification*))
    (write-string "</w:tc>" stream))
  (write-string "</w:tr>" stream))

(defmethod print-tagged-element ((tag (eql '3bmd-grammar::th)) stream rest)
  )

(defmethod print-tagged-element ((tag (eql '3bmd-grammar::td)) stream rest)
  )

;;; code blocks

(defun render-code-block (stream lang params content)
  (declare (ignore params))
  (if (alexandria:emptyp lang)
      (print-verbatim stream (list content))
      (let* ((l (alexandria:make-keyword lang)) ; FIXME: be nicer for unkown lang
	     (runs (recolor:recolor-string content l)))
	(write-string "<w:p><w:pPr><w:pStyle w:val=\"mdcodeblock\"/></w:pPr>" stream)
	(write-string runs stream)
	(write-string "</w:p>" stream))))

(defmethod print-tagged-element ((tag (eql '3bmd-code-blocks::code-block)) stream rest)
  (destructuring-bind (&key lang params content) rest
    (render-code-block stream lang params content)))

;;; print-element

(defgeneric print-element (elem stream))

(defmethod print-element ((elem null) stream)
  "")

(defmethod print-element ((elem (eql :apostrophe)) stream)
  (if *in-code*
      (print-element "'" stream)
      (print-element (string #\right_single_quotation_mark) stream)))

(defmethod print-element ((elem string) stream)
  (cond ((string= #\Newline elem)
	 (open-text stream " ")
	 (write-string " " stream))
	(t
	 (open-text stream elem)
	 (print-escaped elem stream))))

(defmethod print-element ((elem cons) stream)
  (if (symbolp (car elem))
      (print-tagged-element (car elem) stream (cdr elem))
      (error "Unknown cons? ~S" elem)))

(defmethod 3bmd::print-doc-to-stream-using-format (doc stream (format (eql :wml)))
  (let ((*references* (3bmd::extract-refs doc))
	*in-code* *in-block-quote* *in-paragraph* *in-run* *in-text* *run-props* *in-list*
	*para-justification* *new-list* (*very-first-list* t))
    (let ((root
	   (plump:parse
	    (with-output-to-string (s)
	      (dolist (elem doc)
		(print-element elem s))
	      (fresh-line s)))))
      (coalesce-all-adjacent-run-text root)
      (plump:serialize root stream))))

;;;;;;;;;;;;;;;;;;;;;;;;;;;;;

(defparameter *md-heading-sizes* #(30 24 21 18 16 14))

(defun make-md-heading-style (i)
  (wuss:compile-style-to-element
   `(:style type "paragraph" custom-style 1 style-id ,(format nil "mdheading~D" (1+ i))
	    (:name ,(format nil "MD Heading ~D" (1+ i))
	     :ui-priority 1
	     :q-format
	     :p-pr
	     (:keep-lines
	      :keep-next
	      :outline-lvl ,(princ-to-string i)
	      :spacing after ,(princ-to-string (- 240 (* 20 1))))
	     :r-pr
	     (:b
	      :sz ,(princ-to-string (* 2 (aref *md-heading-sizes* i))))))))

(defun make-md-code-style ()
  (wuss:compile-style-to-element
   '(:style type "character" custom-style 1 style-id "mdcode"
     (:name "MD Code"
      :ui-priority 1
      :q-format
      :r-pr
      (:no-proof
       :r-fonts ascii "Consolas"
       :sz 20
       :shd 1 color "auto" fill "DCDCDC")))))

(defun make-md-code-block-style ()
  (wuss:compile-style-to-element
   '(:style type "paragraph" custom-style 1 style-id "mdcodeblock"
     (:name "MD Code Block"
      :ui-priority 2
      :q-format
      :p-pr
      (:shd 1 color "auto" fill "DCDCDC")
      :r-pr
      (:no-proof
       :r-fonts ascii "Consolas"
       :sz 20)))))

(defun make-md-link-style ()
  (wuss:compile-style-to-element
   '(:style type "character" custom-style 1 style-id "mdlink"
     (:name "MD Link"
      :ui-priority 1
      :q-format
      :r-pr
      (:u "single"
       :color "0000FF")))))

(defun make-md-quote-style ()
  (wuss:compile-style-to-element
   '(:style type "paragraph" custom-style 1 style-id "mdquote"
     (:name "MD Link"
      :ui-priority 1
      :q-format
      :p-pr
      (:p-bdr
       (:left "single" sz 24 space 4 color "A9A9A9")
       :shd 1 color "auto" fill "DCDCDC"
       :spacing before 240 after 240)))))

(defun make-md-list-paragraph-style ()
  (wuss:compile-style-to-element
   '(:style type "paragraph" custom-style 1 style-id "mdlistparagraph"
     (:name "MD List Paragraph"
      :ui-priority 1
      :q-format
      :p-pr
      (:ind left 720
       :contextual-spacing)))))

(defun make-md-table-style ()
  (wuss:compile-style-to-element
   '(:style type "table" custom-style 1 style-id "mdtable"
     (:name "MD Table"
      :ui-priority 1
      :q-format
      :p-pr
      (:spacing after 0)
      :tbl-pr
      (:tbl-borders
       (:top "single" sz 4 space 0 color "auto"
	:left "single" sz 4 space 0 color "auto"
	:bottom "single" sz 4 space 0 color "auto"
	:right "single" sz 4 space 0 color "auto"
	:inside-h "single" sz 4 space 0 color "auto"
	:inside-v "single" sz 4 space 0 color "auto")
       :tbl-cell-mar
       (:left w 108 type "dxa"
	:right w 108 type "dxa"))))))

(defparameter *md-code-block-styles*
  '(("mdsymbol" "MD Symbol" "770055")
    ("mdspecial" "MD Special" "FF5000")
    ("mdkeyword" "MD Keyword" "770000")
    ("mdcomment" "MD Comment" "007777")
    ("mdstring" "MD String" "777777")
    ("mdatom" "MD Atom" "314F4F")
    ("mdmacro" "MD Macro" "FF5000")
    ("mdvariable" "MD Variable" "36648B")
    ("mdfunction" "MD Function" "884789")
    ("mdattribute" "MD Attribute" "FF5000")
    ("mdcharacter" "MD Character" "0055AA")
    ("mdsyntaxerror" "MD Syntax Error" "FF0000")
    ("mddiff-deleted" "MD Diff-Deleted" "5F2121")
    ("mddiff-added" "MD Diff-Added" "215F21")))

(defun md-code-block-style (id name color)
  `(:style type "character" custom-style 1 style-id ,id
     (:name ,name
      :ui-priority 2
      :q-format
      :r-pr
      (:color ,color))))

(defun md-code-block-styles ()
  (mapcar (lambda (entry)
	    (wuss:compile-style-to-element
	     (apply #'md-code-block-style entry)))
	  *md-code-block-styles*))

(defvar *md-bullets* #(#\Bullet #\White_Bullet #\Black_Small_Square))

(defvar *md-number-formats* #("decimal" "lowerLetter" "lowerRoman"))

(defparameter *md-numbering-definitions*
  `((:abstract-num abstract-num-id 1
     (:multi-level-type "multilevel"
      ,@(loop for i below 9 appending
	     `(:lvl ilvl ,i
	       (:start 1
		:num-fmt ,(aref *md-number-formats* (mod i 3))	    
		:lvl-text ,(format nil "%~D." (1+ i))
		:lvl-jc "left"
		:p-pr
		(:ind left ,(* (1+ i) 720) hanging 360))))))
    (:abstract-num abstract-num-id 2
     (:multi-level-type "hybrid-multilevel"
      ,@(loop for i below 9 appending
	     `(:lvl ilvl ,i
	       (:start 1
		:num-fmt "bullet"
		:lvl-text ,(aref *md-bullets* (mod i 3))       
		:lvl-jc "left"
		:p-pr
		(:ind left ,(* (1+ i) 720) hanging 360))))))
    (:num num-id 1
     (:abstract-num-id 1))
    (:num num-id 2
     (:abstract-num-id 1
      (:lvl-override ilvl 0
       (:start-override 1))))
    (:num num-id 3
     (:abstract-num-id 2))))

(defun add-md-numbering (numbering-part)
  (let ((numbering (first (plump:get-elements-by-tag-name (opc:xml-root numbering-part) "w:numbering"))))
    (dolist (entry
	      (mapcar #'wuss:compile-style-to-element *md-numbering-definitions*))
      (plump:append-child numbering entry))))

(defparameter *md-styles*
  (list (make-md-heading-style 0)
	(make-md-heading-style 1)
	(make-md-heading-style 2)
	(make-md-heading-style 3)
	(make-md-heading-style 4)
	(make-md-heading-style 5)
	(make-md-code-style)
	(make-md-link-style)
	(make-md-quote-style)
	(make-md-table-style)
	(make-md-list-paragraph-style)
	(make-md-code-block-style)
	))

(defun add-md-styles (document)
  (dolist (style (append *md-styles* (md-code-block-styles))) ; FIXME - make code-block optional?
    (alexandria:if-let (existing-style
			(docxplora:find-style-by-id
			 document
			 (plump:attribute style "w:styleId")))
      (progn
	(docxplora:remove-style document existing-style)
	(docxplora:add-style document style))
      (docxplora:add-style document style))))

(defun coalesce-all-adjacent-run-text (root)
  (let ((runs (plump:get-elements-by-tag-name root "w:r")))
    (serapeum:do-each (run runs)
      (docxplora:coalesce-adjacent-text run))))

(defun md->docx (infile &optional outfile)
  (let* ((3bmd:*smart-quotes* t)
	 (3bmd-tables:*tables* t)
	 (3bmd-code-blocks:*code-blocks* t)
	 (*md-infile* (merge-pathnames infile (user-homedir-pathname)))
	 (outfile (or outfile (merge-pathnames (make-pathname :type "docx") *md-infile*)))
	 (document (docxplora:make-document))
	 (*document* (docxplora:add-main-document document))
	 (*numbering-definitions* (docxplora:ensure-numbering-definitions document))
	 (*style-definitions* (docxplora:add-style-definitions document)))
    (add-md-styles document)
    (add-md-numbering *numbering-definitions*)
    (let* ((xml (plump:parse
		 (with-output-to-string (s)
		  (3bmd:parse-and-print-to-stream *md-infile* s :format :wml))))
	   (mdroot (opc:xml-root *document*))
	   (body (docxplora:get-first-element-by-tag-name mdroot "w:body")))
      (setf (plump:children body) (plump:children xml))
      (docxplora:save-document document outfile))))
