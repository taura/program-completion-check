#import "@preview/touying:0.6.1": *
#import themes.metropolis: *
//#import themes.university: *
//#import themes.aqua: *
//#import themes.dewdrop: *
//#import themes.simple: *
//#import themes.stargazer: *
//#import themes.default: *
//#import "@preview/numbly:0.1.0": numbly

#set text(font: ("Liberation Serif", "TakaoMincho"))
//#set text(font: ("Liberation Serif", "Noto Serif CJK JP"))
#set text(size: 11pt)
#let small(x) = text(size: 9pt)[#x]

#let ao(x) = text(blue)[#x]
#let aka(x) = text(red)[#x]
#let blink(x, y) = ao(link(x, y))
#let indent(b) = grid(columns: (2em, auto), [], b)

/* #images("svg/foo", (2,4,7), start: 3)
-> include image sequence svg/foo_L2.svg, svg/foo_L4.svg, and svg/foo_L7.svg
   at 3, 4, 5 */
#let images(prefix, rng, start: 1, ..kwargs) = for (i, j) in rng.enumerate() [
  #only(i+start, image(prefix + "_L" + str(j) + ".svg", ..kwargs))
]

#show: metropolis-theme.with(
    //aspect-ratio: "16-9",
    //footer:  self => [#small[跡見学園女子大学FD]],
  config-info(
    title: [スライドテンプレート],
    author: [田浦 健次朗],
    date: [2024/05/09],
  ),
)

#set text(size: 32pt)
//#title-slide()

//#outline(depth: 1)
== これは何?

#image("svg/dataflow.svg", width: 100%)

== 科目コード・学籍番号の正規化

- `123４５６78` #ao[`        # 4, 5, 6が全角orz`] \
  $->$ `12345678` #ao[`       # 全角文字は半角に`] \
  $->$ `000000012345678` #ao[`# 全て数字-> 0を足して15桁に`]

- `FEN-５６78` \
  $->$ `FEN-5678` #ao[`        # 全角文字は半角に`] \
  (全て数字ではないのでここまで)

