#!/bin/bash
#src=overview.typ
src=overview-slide.typ
prefix=$(basename ${src} .typ)
#dst=~/0main/${prefix}.pdf
#dst=../docs/slides/${prefix}.pdf
dst=./${prefix}.pdf
typst watch --open evince ${src} ${dst}
