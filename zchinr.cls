% File: zchinr.cls 
% Copyright 2024 Jasper Habicht (mail(at)jasperhabicht.de).
% 
% This work may be distributed and/or modified under the
% conditions of the LaTeX Project Public License version 1.3c,
% available at http://www.latex-project.org/lppl/.
% 
% This file is part of the `zchinr' package (The Work in LPPL)
% and all files in that bundle must be distributed together.
% 
% This work has the LPPL maintenance status `maintained'.
% 
\ProvidesExplClass{zchinr}
  {2024/08/14} {2.0} {Document class for ZChinR}

\bool_new:N \l__zchinr_prepress_bool
\bool_new:N \l__zchinr_standalone_bool

\keys_define:nn { zchinr / global } { 
  prepress   .bool_set:N = \l__zchinr_prepress_bool ,
  prepress   .default:n  = { true } ,
  prepress   .initial:n  = { false } ,
  standalone .bool_set:N = \l__zchinr_standalone_bool ,
  standalone .default:n  = { true } ,
  standalone .initial:n  = { false } 
}

\ProcessKeyOptions [ zchinr / global ]
\LoadClass [ a4paper , twoside , 10pt ] { report }

\msg_new:nnn { zchinr } { no-color-profile } {
  \msg_warning_text:n { zhinr } \iow_newline:
  Color ~ profile ~ not ~ defined ~ in ~ preamble.
}

\msg_new:nnn { zchinr } { pagination-unknown } {
  \msg_warning_text:n { zhinr } \iow_newline:
  Ignoring ~ unknown ~ pagination ~ setting.
}

\msg_new:nnn { zchinr } { toc-style-unknown } {
  \msg_warning_text:n { zhinr } \iow_newline:
  Ignoring ~ unknown ~ toc ~ style ~ setting.
}

\tl_new:N \l_zchinr_color_profile_name_tl
\tl_new:N \l_zchinr_color_profile_info_tl
\tl_new:N \l_zchinr_color_profile_path_tl

\tl_new:N \l_zchinr_issue_year_tl
\tl_new:N \l_zchinr_issue_number_tl
\tl_new:N \l_zchinr_issue_issn_tl

\tl_new:N \l_zchinr_article_first_page_tl
\tl_new:N \l_zchinr_article_doi_tl
\tl_new:N \l_zchinr_article_licence_tl
\tl_new:N \l_zchinr_article_licence_url_tl
\tl_new:N \l_zchinr_article_author_tl
\tl_new:N \l_zchinr_article_author_short_tl
\tl_new:N \l_zchinr_article_author_toc_tl
\tl_new:N \l_zchinr_article_author_copy_tl
\tl_new:N \l_zchinr_article_author_note_tl
\tl_new:N \l_zchinr_article_title_tl
\tl_new:N \l_zchinr_article_title_pdf_tl
\tl_new:N \l_zchinr_article_title_alt_tl
\tl_new:N \l_zchinr_article_title_short_tl
\tl_new:N \l_zchinr_article_title_toc_tl
\tl_new:N \l_zchinr_article_abstract_tl
\tl_new:N \l_zchinr_article_abstract_alt_tl
\tl_new:N \l_zchinr_article_category_tl
\tl_new:N \l_zchinr_article_toc_style_tl
\str_new:N \l_zchinr_article_pagination_str
\bool_new:N \l_zchinr_article_twocolumn_bool
\bool_new:N \l_zchinr_article_page_break_bool
\bool_new:N \l_zchinr_article_small_title_bool
\clist_new:N \l_zchinr_article_title_style_clist

\dim_new:N \l_zchinr_title_distance_dim

\tl_new:N \l_zchinr_local_zchir_tl
\tl_new:N \l_zchinr_local_abstract_tl
\tl_new:N \l_zchinr_local_issn_tl
\tl_new:N \l_zchinr_local_doi_tl
\tl_new:N \l_zchinr_local_licence_tl

\keys_define:nn { zchinr / setup } { 
  color ~ profile          .code:n     = { \keys_set:nn { zchinr / setup / color ~ profile } {#1} } ,
  color ~ profile / name   .tl_set:N   = \l_zchinr_color_profile_name_tl ,
  color ~ profile / name   .default:n  = { Custom } ,
  color ~ profile / info   .tl_set:N   = \l_zchinr_color_profile_info_tl ,
  color ~ profile / info   .default:n  = { none } ,
  color ~ profile / path   .tl_set:N   = \l_zchinr_color_profile_path_tl ,
  title ~ distance         .dim_set:N   = \l_zchinr_title_distance_dim ,
  title ~ distance         .initial:n   = { 35mm } ,
  localization             .code:n      = { \keys_set:nn { zchinr / setup / localization } {#1} } ,
  localization / zchinr    .tl_set:N    = \l_zchinr_local_zchinr_tl ,
  localization / zchinr    .initial:n   = { ZChinR } ,
  localization / abstract  .tl_set:N    = \l_zchinr_local_abstract_tl ,
  localization / abstract  .initial:n   = { Abstract } ,
  localization / issn      .tl_set:N    = \l_zchinr_local_issn_tl ,
  localization / issn      .initial:n   = { ISSN } ,
  localization / doi       .tl_set:N    = \l_zchinr_local_doi_tl ,
  localization / doi       .initial:n   = { DOI } ,
  localization / licence   .tl_set:N    = \l_zchinr_local_licence_tl ,
  localization / licence   .initial:n   = { Lizenz } ,
  issue                    .code:n      = { \keys_set:nn { zchinr / setup / issue } {#1} } ,
  issue / year             .tl_set:N    = \l_zchinr_issue_year_tl , 
  issue / number           .tl_set:N    = \l_zchinr_issue_number_tl , 
  issue / issn             .tl_set:N    = \l_zchinr_issue_issn_tl ,
  issue / issn             .initial:n   = { 1613-5768, ~ online ~ 2366-7125 } ,
  article                  .code:n      = { \keys_set:nn { zchinr / setup / article } {#1} } ,
  article / pagination     .str_set:N   = \l_zchinr_article_pagination_str ,
  article / pagination     .initial:n   = { arabic } , % arabic , roman , none
  article / first ~ page   .tl_set:N    = \l_zchinr_article_first_page_tl , 
  article / doi            .tl_set:N    = \l_zchinr_article_doi_tl , 
  article / licence        .tl_set:N    = \l_zchinr_article_licence_tl ,
  article / licence        .initial:n   = { CC ~ BY ~ 4.0 } ,
  article / licence ~ url  .tl_set:N    = \l_zchinr_article_licence_url_tl ,
  article / licence ~ url  .initial:n   = { https://creativecommons.org/licenses/by/4.0/ } ,
  article / author         .tl_set:N    = \l_zchinr_article_author_tl , 
  article / author ~ short .tl_set:N    = \l_zchinr_article_author_short_tl , 
  article / author ~ short .initial:n   = { \l_zchinr_article_author_tl } , 
  article / author ~ toc   .tl_set:N    = \l_zchinr_article_author_toc_tl , 
  article / author ~ toc   .initial:n   = { \l_zchinr_article_author_tl } , 
  article / author ~ copy  .tl_set:N    = \l_zchinr_article_author_copy_tl , 
  article / author ~ copy  .initial:n   = { \l_zchinr_article_author_tl } , 
  article / author ~ note  .tl_set:N    = \l_zchinr_article_author_note_tl , 
  article / title          .tl_set:N    = \l_zchinr_article_title_tl , 
  article / title ~ pdf    .tl_set:N    = \l_zchinr_article_title_pdf_tl , 
  article / title ~ pdf    .initial:n   = { \l_zchinr_article_title_tl } ,
  article / title ~ alt    .tl_set:N    = \l_zchinr_article_title_alt_tl , 
  article / title ~ short  .tl_set:N    = \l_zchinr_article_title_short_tl , 
  article / title ~ short  .initial:n   = { \l_zchinr_article_title_tl } ,
  article / title ~ toc    .tl_set:N    = \l_zchinr_article_title_toc_tl , 
  article / title ~ toc    .initial:n   = { \l_zchinr_article_title_tl } ,
  article / abstract       .tl_set:N    = \l_zchinr_article_abstract_tl , 
  article / abstract ~ alt .tl_set:N    = \l_zchinr_article_abstract_alt_tl ,
  article / category       .tl_set:N    = \l_zchinr_article_category_tl , 
  article / category       .initial:n   = { Aufsätze } , 
  article / two ~ columns  .bool_set:N  = \l_zchinr_article_twocolumn_bool , 
  article / two ~ columns  .default:n   = { true } ,
  article / two ~ columns  .initial:n   = { false } ,
  article / page ~ break   .bool_set:N  = \l_zchinr_article_page_break_bool , 
  article / page ~ break   .default:n   = { true } ,
  article / page ~ break   .initial:n   = { true } ,
  article / small ~ title  .bool_set:N  = \l_zchinr_article_small_title_bool , 
  article / small ~ title  .default:n   = { true } ,
  article / small ~ title  .initial:n   = { false } ,
  article / title ~ style  .clist_set:N = \l_zchinr_article_title_style_clist , 
  article / title ~ style  .initial:n   = { title } , % category , title , author , abstract , phantom
  article / toc ~ style    .tl_set:N    = \l_zchinr_article_toc_style_tl , 
  article / toc ~ style    .initial:n   = { author ~ first } , % author first , author last , no author
  reset                    .code:n      = {
    \keys_set:nn { zchinr / setup / article } {
      pagination = { arabic } ,
      first ~ page = { } ,
      licence = { CC ~ BY ~ 4.0 } ,
      licence ~ url = { https://creativecommons.org/licenses/by/4.0/ } ,
      author = { } ,
      author ~ short = { \l_zchinr_article_author_tl } ,
      author ~ toc = { \l_zchinr_article_author_tl } ,
      author ~ copy = { \l_zchinr_article_author_tl } , 
      author ~ note = { } ,
      title = { } ,
      title ~ pdf = { \l_zchinr_article_title_tl } ,
      title ~ alt = { } ,
      title ~ short = { \l_zchinr_article_title_tl } ,
      title ~ toc = { \l_zchinr_article_title_tl } ,
      abstract = { } , 
      abstract ~ alt = { } , 
      two ~ columns = { false } ,
      page ~ break = { true } ,
      small ~ title = { false } ,
      title ~ style = { title } ,
      toc ~ style = { author ~ first }
    }
  }
}

\NewDocumentCommand { \zchinrsetup } { m } {
  \keys_set:nn { zchinr / setup } {#1}
}

\tl_new:N \l__zchinr_header_mark_even_tl
\tl_new:N \l__zchinr_header_mark_odd_tl
\tl_new:N \l__zchinr_header_mark_page_tl

\tl_new:N \l__zchinr_toc_entry_tl

\dim_new:N \l__zchinr_tabindent_dim

\cs_new:Npn \__zchinr_halfquad: { 
  \hskip .5em \relax
}

\cs_new:Npn \__zchinr_hfillrule: {
  \leavevmode\leaders\hrule height .75pt\hfill\kern\z@
}

\RequirePackage[english, german]{babel}

\RequirePackage{fontspec}
\setmainfont{TeX ~ Gyre ~ Pagella}
\setsansfont{TeX ~ Gyre ~ Heros ~ Cn}
\newfontfamily\zchirsansfamily[Mapping=tex-text, LetterSpace=2]{TeX ~ Gyre ~ Heros}

\RequirePackage{xeCJK} 
\XeTeXlinebreaklocale "zh" 
\XeTeXlinebreakskip = 0pt plus 1pt minus .5pt
\setCJKmainfont{Noto ~ Serif ~ CJK ~ SC}
\setCJKsansfont{Noto ~ Sans ~ CJK ~ SC}

\RequirePackage{unicode-math}
\setmathfont{TeX ~ Gyre ~ Pagella ~ Math}

\NewDocumentCommand{\zhs}{ m }{\textup{#1}}
\NewDocumentCommand{\zht}{ m }{\textup{#1}}

\RequirePackage[
  inner=25mm,
  outer=30mm,
  top=30mm,
  bottom=25mm,
  headsep=10mm,
  footskip=5mm
]{geometry}
\setlength{\columnsep}{7.5mm}

\cs_new:Npn \__zchinr_cropmark_nw: {
  \begin{picture}(0,0)
    \thinlines\unitlength1pt\linethickness{0.25pt}
    \put(-6,0){\line(-1,0){20}}
    \put(0,6){\line(0,1){20}}
  \end{picture}
}
\cs_new:Npn \__zchinr_cropmark_ne: {
  \begin{picture}(0,0)
    \thinlines\unitlength1pt\linethickness{0.25pt}
    \put(6,0){\line(1,0){20}}
    \put(0,6){\line(0,1){20}}
  \end{picture}
}
\cs_new:Npn \__zchinr_cropmark_sw: {
  \begin{picture}(0,0)
    \thinlines\unitlength1pt\linethickness{0.25pt}
    \put(-6,0){\line(-1,0){20}}
    \put(0,-6){\line(0,-1){20}}
  \end{picture}
}
\cs_new:Npn \__zchinr_cropmark_se: {
  \begin{picture}(0,0)
    \thinlines\unitlength1pt\linethickness{0.25pt}
    \put(6,0){\line(1,0){20}}
    \put(0,-6){\line(0,-1){20}}
  \end{picture}
}

\bool_if:NT \l__zchinr_prepress_bool {
  \geometry{
    layout=a4paper,
    paperwidth=230mm,
    paperheight=317mm,
    layouthoffset=10mm,
    layoutvoffset=10mm
  }

  \AddToHook{shipout/background}{
    \put(10mm,-10mm){\__zchinr_cropmark_nw:}
    \put(\paperwidth-10mm,-10mm){\__zchinr_cropmark_ne:}
    \put(10mm,10mm-\paperheight){\__zchinr_cropmark_sw:}
    \put(\paperwidth-10mm,10mm-\paperheight){\__zchinr_cropmark_se:}

    \special{pdf:put ~ @thispage ~ <<
      /MediaBox ~ [0.0000 ~ 0.0000 ~ 651.9690 ~ 898.5830]
      /CropBox ~ [0.0000 ~ 0.0000 ~ 651.9690 ~ 898.5830]
      /BleedBox ~ [19.8425 ~ 19.8425 ~ 632.1265 ~ 878.7405]
      /TrimBox ~ [28.3465 ~ 28.3465 ~ 623.6225 ~ 870.2365]
    >>}
  }

  \AtBeginDocument{
    \special{pdf:minorversion ~ 3}
    \special{pdf:docinfo ~ <<
      /GTS_PDFXVersion ~ (PDF/X-1:2001)
      /GTS_PDFXConformance ~ (PDF/X-1a:2001)
    >>}
    \str_if_empty:NTF { \l_zchinr_color_profile_path_tl } {
      \msg_warning:nn { zchinr } { no-color-profile } 
    } {
      \special{pdf:fstream ~ @cmykdata ~ (\l_zchinr_color_profile_path_tl) ~ <<
        /N ~ 4^^J/Alternate/DeviceCMYK
      >>}
    }
    \str_if_empty:NTF { \l_zchinr_color_profile_path_tl } {
      \special{pdf:put ~ @catalog ~<<
        /PageMode ~ /UseNone
        /OutputIntents ~ [<<
          /Type ~ /OutputIntent
          /S ~ /GTS_PDFX
          /OutputConditionIdentifier ~ (Custom)
          /Info ~ (none)
          /RegistryName ~ (http://www.color.org/)
        >>]
      >>}
    } {
      \special{pdf:put ~ @catalog ~ <<
        /PageMode ~ /UseNone
        /OutputIntents ~ [<<
          /Type ~ /OutputIntent
          /S ~ /GTS_PDFX
          /DestOutputProfile ~ @cmykdata
          /OutputConditionIdentifier ~ (\l_zchinr_color_profile_id_tl)
          /Info ~ (\l_zchinr_color_profile_info_tl)
          /RegistryName ~ (http://www.color.org/)
        >>]
      >>}
    }
  }
}

\RequirePackage{afterpage}
\RequirePackage{array}
\RequirePackage{longtable}
\RequirePackage{booktabs}
\RequirePackage{enumitem}
\RequirePackage{titletoc}
\RequirePackage[pagestyles]{titlesec}

\frenchspacing
\clubpenalty=10000
\widowpenalty=10000
\doublehyphendemerits=5000000

\setlength{\parindent}{15pt}
\setlength{\parskip}{2pt plus 1pt minus 1pt}

\setlength{\arrayrulewidth}{.75pt}

% Style figure and table captions.
\RequirePackage[
  labelfont=bf,
  labelsep=period,
]{caption}

% Set up hyperreferences
\PassOptionsToPackage{hyphens}{url}
\bool_if:NTF \l__zchinr_prepress_bool {
  \RequirePackage[bookmarks=false]{hyperref}
  \RequirePackage{datetime}
  \AtBeginDocument{
    \NoHyper
    \hypersetup{
      pdftitle={\l_zchinr_local_zchinr_tl{} ~ \l_zchinr_issue_year_tl{}},
      pdfinfo={
        ModDate={D:\pdfdate},
        Trapped={False},      
      }
    }
  }
} {
  \RequirePackage[hidelinks]{hyperref}
}

% Style header and footer. Header stays empty while footer will contain journal information and page number. 
\newpagestyle{zchinrdefault}{
  \sethead
    [\l__zchinr_header_mark_page_tl]
    [\l__zchinr_header_mark_odd_tl]
    [\l_zchinr_local_zchinr_tl{} ~ \l_zchinr_issue_year_tl]
    {\l_zchinr_local_zchinr_tl{} ~ \l_zchinr_issue_year_tl}
    {\l__zchinr_header_mark_even_tl}
    {\l__zchinr_header_mark_page_tl}
  \setfoot{}{}{}
}
\pagestyle{zchinrdefault}

\newpagestyle{zchinrtitle}{
  \sethead{}{}{}
  \setfoot
    {{\parbox[b]{0.4\textwidth}{\raggedright
      \l_zchinr_local_issn_tl \c_colon_str{} ~ \l_zchinr_issue_issn_tl \\
      \l_zchinr_local_doi_tl \c_colon_str{} ~ 
        \href{https://doi.org/\l_zchinr_article_doi_tl}{\l_zchinr_article_doi_tl}
    }}}
    {\l__zchinr_header_mark_page_tl}
    {{\parbox[b]{0.4\textwidth}{\raggedleft
      \copyright{} ~ \l_zchinr_issue_year_tl{} ~ \l_zchinr_article_author_tl \\ 
      \l_zchinr_local_licence_tl \c_colon_str{} ~ 
        \href{\l_zchinr_article_licence_url_tl}{\l_zchinr_article_licence_tl}
    }}}
}

\newpagestyle{zchirempty}{
  \sethead{}{}{}
  \setfoot{}{}{}
}

% Style section titles. 
\titleformat{\section}{\fontsize{11}{13}\selectfont\bfseries}{}{0pt}{}
\titleformat{\subsection}{\fontsize{11}{13}\selectfont\bfseries}{}{0pt}{}
\titleformat{\subsubsection}{\fontsize{11}{13}\selectfont\bfseries}{}{0pt}{}
\titlespacing*{\section}{0pt}{5pt plus 1pt}{5pt plus 1pt}
\titlespacing*{\subsection}{0pt}{5pt plus 1pt}{5pt plus 1pt}
\titlespacing*{\subsubsection}{0pt}{5pt plus 1pt}{5pt plus 1pt}

\titleformat{\paragraph}[runin]{\bfseries}{}{2em}{\noindent}
\titleformat{\subparagraph}[runin]{\bfseries}{}{2em}{\noindent}
\titlespacing*{\paragraph}{0pt}{5pt plus 1pt}{1em plus 5pt minus 10pt}
\titlespacing*{\subparagraph}{0pt}{5pt plus 1pt}{1em plus 5pt minus 10pt}

\titlecontents{section}[0pt]{}{\hspace*{-15pt}}{}
  { \__zchinr_halfquad: \titlerule*[5pt]{.}\contentspage}
\titlecontents{subsection}[15pt]{}{\hspace*{-15pt}}{}
  { \__zchinr_halfquad: \titlerule*[5pt]{.}\contentspage}
\titlecontents{subsubsection}[30pt]{}{\hspace*{-15pt}}{}
  { \__zchinr_halfquad: \titlerule*[5pt]{.}\contentspage}
\titlecontents{paragraph}[45pt]{}{\hspace*{-15pt}}{}
  { \__zchinr_halfquad: \titlerule*[5pt]{.}\contentspage}
\titlecontents{subparagraph}[60pt]{}{\hspace*{-15pt}}{}
  { \__zchinr_halfquad: \titlerule*[5pt]{.}\contentspage}

\titlecontents{part}[0pt]{\bigskip\Large\scshape\bfseries}{}{}{}[\bigskip]
\titlecontents{chapter}[15pt]{}{}{}{\hfill\contentspage}

% Style footnotes.
\cs_set:Npn \@makefntext #1 {
  \parbox[t]{2em}{\@thefnmark}
  \parbox[t]{\dimexpr\columnwidth-2em}{#1}
}
\tl_set:Nn \footnoterule {}
\setlength{\skip\footins}{15pt plus 5pt minus 5pt}

% Create listing environment for references. 
\newlist{references}{enumerate}{1}
\setlist[references]{label*=[\arabic*]}

% Style URLs.
\urlstyle{same}
\DeclareUrlCommand\Hurl{\def\UrlLeft{<\,}\def\UrlRight{\,>}}
\tl_put_right:Nn \UrlBreaks {\do\^\do\_}

% Create command for automatically linked e-mail addresses. 
\NewDocumentCommand { \email } { m } {
  \href{mailto:#1}{\protect\nolinkurl{#1}}
}

% Create documentation environment. 
\NewDocumentEnvironment { documentation } { } {
  \tl_set:Nn \arraystretch { 1.25 }
  \noindent\begin{longtable}{
    @{} 
      >{ \dim_set:Nn \l__zchinr_tabindent_dim { 2em } 
        \hskip \l__zchinr_tabindent_dim } 
      p{54mm}
    @{\hskip\columnsep} 
      >{ \dim_set:Nn \l__zchinr_tabindent_dim { 15pt } 
        \hskip \l__zchinr_tabindent_dim } 
      p{\dimexpr\textwidth - 54mm - \columnsep}
    @{}
  }
} {
  \end{longtable}
}

\NewDocumentCommand { \documentationheader } { O{c} O{} m o o } {
  \leavevmode
  \IfNoValueF{#4}{
    \phantomsection
    \addcontentsline{toc}{#4}{\IfNoValueTF{#5}{#3}{#5}}
  }
  \hskip -\l__zchinr_tabindent_dim 
  \rule{0pt}{\dimexpr\baselineskip+5pt}
  \str_if_eq:nnT {#1} { c } { \centering }
  \str_if_eq:nnT {#1} { l } { \raggedright }
  \str_if_eq:nnT {#1} { r } { \raggedleft }
  \str_if_eq:nnF {#2} { n } { \bfseries } 
  \arraybackslash #3 
  \rule[-10pt]{0pt}{0pt}
}

% Create addresses environment.
\NewDocumentEnvironment { addresses } { } {
  \tl_set:Nn \arraystretch { 1.75 }
  \noindent\begin{longtable}{
    @{} 
      p{\dimexpr.5\textwidth - .5\columnsep}
    @{\hskip\columnsep} 
      p{\dimexpr.5\textwidth - .5\columnsep}
    @{}
  }
} {
  \end{longtable}
}

\NewDocumentCommand { \addressessection } { m o } {
  \phantomsection
  \addcontentsline{toc}{chapter}{\IfNoValueTF{#2}{#1}{#2}}
  \noindent{\huge #1}
}

\NewDocumentCommand { \addressesheader } { m o o } {
  \leavevmode
  \IfNoValueF{#2}{
    \phantomsection
    \addcontentsline{toc}{#2}{\IfNoValueTF{#3}{#1}{#3}}
  }
  \rule{0pt}{\dimexpr\baselineskip+15pt}
  \raggedright\bfseries\arraybackslash #1
}

\NewExpandableDocumentCommand { \addressesspan } { m } {
  \multicolumn{2}{@{}p{\textwidth}@{}}{#1}
}

\NewDocumentCommand { \addressessep } { m } {
  \char"00B7{}
}

% Create imprint environment.
\NewDocumentEnvironment { imprint } { } {
  \tl_set:Nn \arraystretch { 1.25 }
  \noindent\begin{longtable}{
    @{} p{30mm}
    @{\hskip\columnsep} p{\dimexpr\textwidth - 30mm - 35mm - 1.5\columnsep}
    @{\hskip .5\columnsep} >{\raggedleft\arraybackslash} p{35mm}
    @{}
  }
} {
  \end{longtable}
}

% Create commands for TOC. 
\NewDocumentCommand { \tocentry } { m }{
  \bigskip
  \noindent{\Large\scshape\bfseries \text_lowercase:n {#1} }
}

% Create commands for TOC. 
\NewDocumentCommand { \printissuetoc } { } {
  \startcontents
  \printcontents{}{-1}[0]{}
}

\newlist{toc}{itemize}{1}
\setlist[toc]{leftmargin=15mm, label*={}}

% Redefine abstract environment.
\RenewDocumentEnvironment { abstract } { O{\textwidth} } {
  \begin{minipage}[t]{#1}
  \setlength{\parindent}{15pt}
  \noindent
} {
  \end{minipage}
}

% Create subcontents environment.
\NewDocumentEnvironment { subcontents } { } {
  \list{}{\leftmargin=35mm \rightmargin=20mm}
  \item\relax
} {
  \endlist
}

\NewDocumentCommand { \printtitle } { } {
  \bool_if:NTF \l_zchinr_article_page_break_bool {
    \clearpage
    \setcounter{footnote}{0}
    \startcontents[zchinrsub]
  } {
    \bigskip
  }
  \thispagestyle{zchinrtitle}

  \tl_if_empty:NF \l_zchinr_article_first_page_tl {
    \setcounter{page}{ \l_zchinr_article_first_page_tl }
  }

  \str_case_e:nnF { \l_zchinr_article_pagination_str } {
    { none } {
      \tl_set:Nn \l__zchinr_header_mark_page_tl { }
    }
    { roman } {
      \tl_set:Nn \l__zchinr_header_mark_page_tl { \roman{page} }
    }
    { arabic } {
      \tl_set:Nn \l__zchinr_header_mark_page_tl { \arabic{page} } 
    }
  } {
    \msg_warning:nn { zchinr } { pagination-unknown } 
  }

  \tl_if_empty:NTF \l_zchinr_article_title_short_tl {
    \tl_set_eq:NN \l__zchinr_header_mark_even_tl \l_zchinr_article_category_tl
    \tl_set_eq:NN \l__zchinr_header_mark_odd_tl \l_zchinr_article_category_tl  
  } {
    \tl_if_empty:NTF \l_zchinr_article_author_short_tl {
      \tl_set_eq:NN \l__zchinr_header_mark_even_tl \l_zchinr_article_title_short_tl
      \tl_set_eq:NN \l__zchinr_header_mark_odd_tl \l_zchinr_article_title_short_tl
    } {
      \tl_set:Nn \l__zchinr_header_mark_even_tl {
        { \itshape \l_zchinr_article_author_short_tl , } ~ \l_zchinr_article_title_short_tl
      }
      \tl_set:Nn \l__zchinr_header_mark_odd_tl {
        { \itshape \l_zchinr_article_author_short_tl , } ~ \l_zchinr_article_title_short_tl
      }
    }
  }

  \tl_clear:N \l__zchinr_toc_entry_tl
  \str_case_e:nnF { \l_zchinr_article_toc_style_tl } {
    { author ~ first } {
      \tl_set:Nn \l__zchinr_toc_entry_tl { 
        { \itshape \l_zchinr_article_author_toc_tl , } ~ 
          \l_zchinr_article_title_toc_tl
      }
    }
    { author ~ last } {
      \tl_set:Nn \l__zchinr_toc_entry_tl { 
        \l_zchinr_article_title_toc_tl \newline 
          { \itshape ( \l_zchinr_article_author_toc_tl ) }
      }
    }
    { no ~ author } {
      \tl_set:Nn \l__zchinr_toc_entry_tl { 
        \l_zchinr_article_title_toc_tl 
      }
    }
  } {
    \msg_warning:nn { zchinr } { toc-style-unknown }
  }

  \tl_set:Nn \thefootnote { \char"002A }
  \bool_if:NTF \l_zchinr_article_page_break_bool {
    \bool_if:NTF \l_zchinr_article_twocolumn_bool {
      \twocolumn[
        \__zchir_article_title:
      ]
      \sloppy\flushbottom
      \enlargethispage{-10mm}
      \afterpage{\enlargethispage{-10mm}}
    } {
      \onecolumn
      \__zchir_article_title:
      \enlargethispage{-10mm}
    }
  } {
    \__zchir_article_title:
    \enlargethispage{-10mm}
    \bool_if:NT \l_zchinr_article_twocolumn_bool {
      \afterpage{\enlargethispage{-10mm}}
    }  
  }
  \tl_if_empty:NF \l_zchinr_article_author_note_tl {
    \footnotetext{\l_zchinr_article_author_note_tl}
    \addtocounter{footnote}{-1}
  }
  \tl_set:Nn \thefootnote { \arabic{footnote} }
  \@afterindentfalse\@afterheading
}

\cs_new:Npn \__zchir_article_author: {
  \l_zchinr_article_author_tl
  \tl_if_empty:NF \l_zchinr_article_author_note_tl {
    \, \footnotemark
  }
}

\cs_new:Npn \__zchir_article_title: {
  \clist_if_in:NnT \l_zchinr_article_title_style_clist { category } {
    \bool_if:NF \l__zchinr_standalone_bool {
      \phantomsection
      \addcontentsline{toc}{part}{
        \texorpdfstring{ \text_lowercase:n { \l_zchinr_article_category_tl } }
          {\l_zchinr_article_category_tl}
      }
    }
    \noindent{
      \int_if_even:nTF { \value{page} } {
        \raggedright
      } {
        \raggedleft
      }
      \LARGE\scshape\bfseries \text_lowercase:n { \l_zchinr_article_category_tl }
    } \__zchinr_halfquad: \__zchinr_hfillrule:
    \par\vskip 7mm
  }
  
  \clist_if_in:NnTF \l_zchinr_article_title_style_clist { title } {
    \bool_if:NF \l__zchinr_standalone_bool {
      \phantomsection
      \addcontentsline{toc}{chapter}{
        \texorpdfstring{\l__zchinr_toc_entry_tl}{\l_zchinr_article_title_tl}
      }
    }
    \bool_if:NTF \l_zchinr_article_small_title_bool {
      \noindent{\fontsize{11}{13}\selectfont\bfseries\l_zchinr_article_title_tl}
      \par\bigskip\nopagebreak
      \noindent{\itshape \__zchir_article_author: }
      \par\bigskip\nopagebreak
    } {
      \noindent\parbox[t][\l_zchinr_title_distance_dim][t]{\textwidth}{
        \raggedright{\huge\l_zchinr_article_title_tl}
        \clist_if_in:NnT \l_zchinr_article_title_style_clist { author } {
          \par\bigskip
          {\fontsize{11}{13}\selectfont\itshape \__zchir_article_author: }
        }
      }
      \par\vskip 4mm\bigskip
    }
  } {
    \clist_if_in:NnT \l_zchinr_article_title_style_clist { phantom } {
      \phantomsection
      \addcontentsline{toc}{chapter}{
        \texorpdfstring{\l__zchinr_toc_entry_tl}{\l_zchinr_article_title_tl}
      }
    }
  }

  \clist_if_in:NnT \l_zchinr_article_title_style_clist { abstract } {
    \noindent{\fontsize{11}{13}\selectfont\itshape\bfseries\l_zchinr_local_abstract_tl}
    \par\vskip 8pt
    \str_if_empty:NTF \l_zchinr_article_abstract_alt_tl {
      \noindent\begin{abstract}
        \itshape\l_zchinr_article_abstract_tl
      \end{abstract}
    } {
      \noindent\begin{abstract}[\columnwidth]
        \itshape\l_zchinr_article_abstract_tl
      \end{abstract}
      \hskip\columnsep
      \begin{abstract}[\columnwidth]
        \itshape{\bfseries\l_zchinr_article_title_alt_tl} ~ --- ~
        \l_zchinr_article_abstract_alt_tl
      \end{abstract}
    } 
    \par\vskip 12mm
  }

  \clist_if_in:NnT \l_zchinr_article_title_style_clist { contents } {
    \begin{subcontents}
      \printcontents[zchinrsub]{}{1}{}
    \end{subcontents}
    \par\vskip 12mm
  }
}

\NewDocumentCommand { \inputpart } { s m } {
  \IfBooleanTF{#1}{
    \bool_if:NF \l__zchnr_standalone_bool {
      \input{#2}
    }
  }{
    \input{#2}
  }
}

% EOF