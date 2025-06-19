(defun xml-escape (str)
  (setq str (vl-string-subst "&amp;" "&" str))
  (setq str (vl-string-subst "&lt;" "<" str))
  (setq str (vl-string-subst "&gt;" ">" str))
  (setq str (vl-string-subst "&quot;" "\"" str))
  str
)

(defun c:tst25 ( / doc ms ent blkName minpt maxpt minList maxList txtEnt txtPt ptList textStr blkFile blkFolder outFile outHandle blkPattern blocksFound )

  ;; �������� ���� � ����� � ��� ����� �� ���������� ���������
  (setq blkFolder (getenv "IZBLOCKPATH"))
  (setq blkPattern (getenv "IZBLOCKNAME"))
  (if (or (not blkFolder) (not blkPattern))
    (progn (prompt "\n��������� ���������� IZBLOCKPATH � IZBLOCKNAME!") (princ) (exit))
  )

  ;; ��� �������� ����� DWG ��� ���� � ����������
  (setq doc (vla-get-ActiveDocument (vlax-get-acad-object)))
  (setq ms  (vla-get-ModelSpace doc))
  (setq blkFile (vl-filename-base (getvar "DWGNAME")))
  (setq outFile (strcat blkFolder "\\IZBLOCK_" blkFile ".xml"))
  (setq outHandle (open outFile "w"))
  (if (not outHandle)
    (progn (prompt (strcat "\n�� ������� ������� ���� ��� ������: " outFile)) (princ) (exit))
  )

  ;; ���������� ��������� XML � ������������ ���
  (write-line "<?xml version=\"1.0\" encoding=\"utf-8\"?>" outHandle)
  (write-line "<blocks>" outHandle)

  (setq blocksFound 0)

  ;; �������� �� ���� ������
  (vlax-for ent ms
    (if (and (= (vla-get-ObjectName ent) "AcDbBlockReference")
             (wcmatch (strcase (vla-get-EffectiveName ent)) (strcat (strcase blkPattern) "*")))
      (progn
        (setq blocksFound (+ blocksFound 1))
        (setq blkName (vla-get-EffectiveName ent))
        (setq minpt (vlax-3d-point '(0 0 0)))
        (setq maxpt (vlax-3d-point '(0 0 0)))
        (vla-GetBoundingBox ent 'minpt 'maxpt)

        ;; �������� ������ ���������
        (defun safe->list (val)
          (cond
            ((= (type val) 'LIST) val)
            ((vl-catch-all-error-p
              (vl-catch-all-apply 'vlax-safearray->list (list val)))
              val)
            (T (vlax-safearray->list val)))
        )

        (setq minList (safe->list minpt))
        (setq maxList (safe->list maxpt))

        ;; ���������� �������� ����� (������ ���� ���)
        (write-line
          (strcat "\t<block name=\"" (xml-escape blkName)
                  "\" minx=\"" (rtos (nth 0 minList) 2 6)
                  "\" maxx=\"" (rtos (nth 0 maxList) 2 6)
                  "\" miny=\"" (rtos (nth 1 minList) 2 6)
                  "\" maxy=\"" (rtos (nth 1 maxList) 2 6)
                  "\">")
          outHandle
        )

        ;; ������ ��������� ������� � ������������ (������ ��� ����)
        (vlax-for txtEnt ms
          (if (or (= (vla-get-ObjectName txtEnt) "AcDbText")
                  (= (vla-get-ObjectName txtEnt) "AcDbMText"))
            (progn
              (setq txtPt (safe->list (vlax-get txtEnt 'InsertionPoint)))
              (if (and (<= (nth 0 minList) (nth 0 txtPt) (nth 0 maxList))
                       (<= (nth 1 minList) (nth 1 txtPt) (nth 1 maxList)))
                (progn
                  (setq textStr (xml-escape (vlax-get txtEnt 'TextString)))
                  (write-line
                    (strcat "\t\t<text content=\"" textStr
                            "\" x=\"" (rtos (nth 0 txtPt) 2 6)
                            "\" y=\"" (rtos (nth 1 txtPt) 2 6)
                            "\" />")
                    outHandle
                  )
                )
              )
            )
          )
        )

        ;; ��������� ���� (������ ���� ���)
        (write-line "\t</block>" outHandle)
      )
    )
  )

  ;; ��������� ������������ ���
  (write-line "</blocks>" outHandle)

  ;; �������� �������� �����
  (close outHandle)
  (if (> blocksFound 0)
    (prompt (strcat "\nXML ���� ������� �������: " outFile))
    (prompt "\n����� �� �������, ���� ������.")
  )
  (princ)
)
