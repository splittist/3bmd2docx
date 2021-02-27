;;;; 3bmd2docx.asd

(asdf:defsystem #:3bmd2docx
  :description "Extends 3bmd to output WordprocessingML"
  :author "John Q. Splittist <splittist@splittist.com>"
  :license  "MIT"
  :version "0.0.1"
  :serial t
  :depends-on (
	       #:docxplora
	       #:wuss
	       #:recolor
	       #:imagesniff

	       #:3bmd
	       #:3bmd-ext-tables
	       #:3bmd-ext-code-blocks

	       #:alexandria
	       #:serapeum
	       #:split-sequence
	       #:plump)
  :components ((:file "package")
	       (:file "mdprinter")))
