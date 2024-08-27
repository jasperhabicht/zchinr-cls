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

\bool_new:N \l_zchinr_prepress_bool
\bool_new:N \l_zchinr_standalone_bool

\keys_define:nn { zchinr / global } { 
  prepress   .bool_set:N = \l_zchinr_prepress_bool ,
  prepress   .default:n  = { true } ,
  prepress   .initial:n  = { false } ,
  standalone .bool_set:N = \l_zchinr_standalone_bool ,
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

\msg_new:nnn { zchinr } { header-style-unknown } {
  \msg_warning_text:n { zhinr } \iow_newline:
  Ignoring ~ unknown ~ header ~ style ~ setting.
}

\msg_new:nnn { zchinr } { toc-style-unknown } {
  \msg_warning_text:n { zhinr } \iow_newline:
  Ignoring ~ unknown ~ toc ~ style ~ setting.
}

\tl_new:N \l_zchinr_color_profile_name_tl
\tl_new:N \l_zchinr_color_profile_info_tl
\str_new:N \l_zchinr_color_profile_path_str

\tl_new:N \l_zchinr_issue_year_tl
\tl_new:N \l_zchinr_issue_volume_tl
\tl_new:N \l_zchinr_issue_number_tl
\tl_new:N \l_zchinr_issue_issn_tl
\tl_new:N \g_zchinr_issue_first_page_tl
\tl_new:N \g_zchinr_issue_last_page_tl
\tl_new:N \g__zchinr_issue_first_page_aux_tl
\tl_new:N \g__zchinr_issue_last_page_aux_tl
\bool_new:N \g__zchinr_issue_first_page_set_bool

\tl_new:N \l_zchinr_default_licence_tl
\tl_new:N \l_zchinr_default_licence_url_tl

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
\str_new:N \l_zchinr_article_pagination_str
\str_new:N \l_zchinr_article_toc_style_str
\str_new:N \l_zchinr_article_header_style_str
\bool_new:N \l_zchinr_article_twocolumn_bool
\bool_new:N \l_zchinr_article_small_title_bool
\clist_new:N \l_zchinr_article_title_style_clist

\dim_new:N \l_zchinr_title_distance_dim

\tl_new:N \l_zchinr_local_zchinr_tl
\tl_new:N \l_zchinr_local_abstract_tl
\tl_new:N \l_zchinr_local_issn_tl
\tl_new:N \l_zchinr_local_doi_tl
\tl_new:N \l_zchinr_local_licence_tl

\tl_new:N \l__zchinr_header_mark_even_tl
\tl_new:N \l__zchinr_header_mark_odd_tl
\tl_new:N \l__zchinr_header_mark_page_tl

\tl_new:N \l__zchinr_toc_entry_tl

\tl_new:N \l__zchinr_fnmark_tl

\dim_new:N \l__zchinr_tabindent_dim

\keys_define:nn { zchinr / setup } { 
  color ~ profile          .code:n     = { \keys_set:nn { zchinr / setup / color ~ profile } {#1} } ,
  color ~ profile / name   .tl_set:N   = \l_zchinr_color_profile_name_tl ,
  color ~ profile / name   .default:n  = { Custom } ,
  color ~ profile / info   .tl_set:N   = \l_zchinr_color_profile_info_tl ,
  color ~ profile / info   .default:n  = { none } ,
  color ~ profile / path   .str_set:N   = \l_zchinr_color_profile_path_str ,
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
  default / licence        .tl_set:N    = \l_zchinr_default_licence_tl ,
  default / licence        .initial:n   = { CC ~ BY ~ 4.0 } ,
  default / licence ~ url  .tl_set:N    = \l_zchinr_default_licence_url_tl ,
  default / licence ~ url  .initial:n   = { https://creativecommons.org/licenses/by/4.0/ } ,
  issue                    .code:n      = { \keys_set:nn { zchinr / setup / issue } {#1} } ,
  issue / year             .tl_set:N    = \l_zchinr_issue_year_tl , 
  issue / volume           .tl_set:N    = \l_zchinr_issue_volume_tl , 
  issue / number           .tl_set:N    = \l_zchinr_issue_number_tl , 
  issue / issn             .tl_set:N    = \l_zchinr_issue_issn_tl ,
  issue / issn             .initial:n   = { 1613-5768, ~ online ~ 2366-7125 } ,
  article                  .code:n      = { \keys_set:nn { zchinr / setup / article } {#1} } ,
  article / pagination     .str_set:N   = \l_zchinr_article_pagination_str ,
  article / pagination     .initial:n   = { arabic } , % arabic , roman , none
  article / first ~ page   .tl_set:N    = \l_zchinr_article_first_page_tl , 
  article / first ~ page   .initial:n   = { } ,
  article / doi            .tl_set:N    = \l_zchinr_article_doi_tl , 
  article / licence        .tl_set:N    = \l_zchinr_article_licence_tl ,
  article / licence        .initial:n   = { \l_zchinr_default_licence_tl } ,
  article / licence ~ url  .tl_set:N    = \l_zchinr_article_licence_url_tl ,
  article / licence ~ url  .initial:n   = { l_zchinr_default_licence_url_tl } ,
  article / author         .tl_set:N    = \l_zchinr_article_author_tl , 
  article / author         .initial:n   = { } , 
  article / author ~ short .tl_set:N    = \l_zchinr_article_author_short_tl , 
  article / author ~ short .initial:n   = { \l_zchinr_article_author_tl } , 
  article / author ~ toc   .tl_set:N    = \l_zchinr_article_author_toc_tl , 
  article / author ~ toc   .initial:n   = { \l_zchinr_article_author_tl } , 
  article / author ~ copy  .tl_set:N    = \l_zchinr_article_author_copy_tl , 
  article / author ~ copy  .initial:n   = { \l_zchinr_article_author_tl } , 
  article / author ~ note  .tl_set:N    = \l_zchinr_article_author_note_tl , 
  article / author ~ note  .initial:n   = { } , 
  article / title          .tl_set:N    = \l_zchinr_article_title_tl , 
  article / title          .initial:n   = { } , 
  article / title ~ pdf    .tl_set:N    = \l_zchinr_article_title_pdf_tl , 
  article / title ~ pdf    .initial:n   = { \l_zchinr_article_title_tl } ,
  article / title ~ alt    .tl_set:N    = \l_zchinr_article_title_alt_tl , 
  article / title ~ alt    .initial:n   = { } , 
  article / title ~ short  .tl_set:N    = \l_zchinr_article_title_short_tl , 
  article / title ~ short  .initial:n   = { \l_zchinr_article_title_tl } ,
  article / title ~ toc    .tl_set:N    = \l_zchinr_article_title_toc_tl , 
  article / title ~ toc    .initial:n   = { \l_zchinr_article_title_tl } ,
  article / abstract       .tl_set:N    = \l_zchinr_article_abstract_tl , 
  article / abstract       .initial:n   = { } , 
  article / abstract ~ alt .tl_set:N    = \l_zchinr_article_abstract_alt_tl ,
  article / abstract ~ alt .initial:n   = { } ,
  article / category       .tl_set:N    = \l_zchinr_article_category_tl , 
  article / category       .initial:n   = { Aufsätze } , 
  article / two ~ columns  .bool_set:N  = \l_zchinr_article_twocolumn_bool , 
  article / two ~ columns  .default:n   = { true } ,
  article / two ~ columns  .initial:n   = { false } ,
  article / small ~ title  .bool_set:N  = \l_zchinr_article_small_title_bool , 
  article / small ~ title  .default:n   = { true } ,
  article / small ~ title  .initial:n   = { false } ,
  article / title ~ style  .clist_set:N = \l_zchinr_article_title_style_clist , 
  article / title ~ style  .initial:n   = { title } , % category , title , author , abstract , phantom
  article / header ~ style .str_set:N   = \l_zchinr_article_header_style_str , 
  article / header ~ style .initial:n   = { author ~ title } , % author title , title , category
  article / toc ~ style    .str_set:N   = \l_zchinr_article_toc_style_str , 
  article / toc ~ style    .initial:n   = { author ~ first } , % author first , author last , no author
  reset                    .code:n      = {
   \keys_set:nn { zchinr / setup / article } {
      pagination = { arabic } ,
      first ~ page = { } ,
      doi = { } ,
      licence = { \l_zchinr_default_licence_tl } ,
      licence ~ url = { \l_zchinr_default_licence_url_tl } ,
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
      small ~ title = { false } ,
      title ~ style = { title } ,
      header ~ style = { author ~ title } ,
      toc ~ style = { author ~ first }
    }
  }
}

\NewDocumentCommand { \zchinrsetup } { s m } {
  \IfBooleanF{#1}{
    \clearpage
  }
  \keys_set:nn { zchinr / setup } {#2}
}

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
\newfontfamily\zchinrsansfamily[Mapping=tex-text, LetterSpace=2]{TeX ~ Gyre ~ Heros}

\RequirePackage{xeCJK} 
\XeTeXlinebreaklocale "zh" 
\XeTeXlinebreakskip = 0pt plus 1pt minus .5pt
\setCJKmainfont{Noto ~ Serif ~ CJK ~ SC}
\setCJKsansfont{Noto ~ Sans ~ CJK ~ SC}

\RequirePackage{unicode-math}
\setmathfont{TeX ~ Gyre ~ Pagella ~ Math}

\NewDocumentCommand{\zhs}{ m }{\textup{#1}}
\NewDocumentCommand{\zht}{ m }{\textup{#1}}

\RequirePackage{microtype}

\RequirePackage[
  inner=25mm,
  outer=30mm,
  top=30mm,
  bottom=25mm,
  headsep=10mm,
  footskip=5mm
]{geometry}
\setlength{\columnsep}{7.5mm}

\RequirePackage{l3draw}
\cs_new:Npn \__zchinr_cropmark_nw: {
  \draw_begin:
    \bool_set_false:N \l_draw_bb_update_bool
    \draw_linewidth:n { .25pt }
    \draw_transform_shift:n { 10mm , -10mm }
    \draw_path_moveto:n { -5pt , 0pt }
    \draw_path_lineto:n { -20pt , 0pt }
    \draw_path_moveto:n { 0 , 5pt }
    \draw_path_lineto:n { 0 , 20pt }
    \draw_path_use:n { stroke }
  \draw_end:
}
\cs_new:Npn \__zchinr_cropmark_ne: {
  \draw_begin:
    \bool_set_false:N \l_draw_bb_update_bool
    \draw_linewidth:n { .25pt }
    \draw_transform_shift:n { \paperwidth - 10mm , -10mm }
    \draw_path_moveto:n { 5pt , 0pt }
    \draw_path_lineto:n { 20pt , 0pt }
    \draw_path_moveto:n { 0 , 5pt }
    \draw_path_lineto:n { 0 , 20pt }
    \draw_path_use:n { stroke }
  \draw_end:
}
\cs_new:Npn \__zchinr_cropmark_sw: {
  \draw_begin:
    \bool_set_false:N \l_draw_bb_update_bool
    \draw_linewidth:n { .25pt }
    \draw_transform_shift:n { 10mm , 10mm - \paperheight }
    \draw_path_moveto:n { -5pt , 0pt }
    \draw_path_lineto:n { -20pt , 0pt }
    \draw_path_moveto:n { 0 , -5pt }
    \draw_path_lineto:n { 0 , -20pt }
    \draw_path_use:n { stroke }
  \draw_end:
}
\cs_new:Npn \__zchinr_cropmark_se: {
  \draw_begin:
    \bool_set_false:N \l_draw_bb_update_bool
    \draw_linewidth:n { .25pt }
    \draw_transform_shift:n { \paperwidth - 10mm , 10mm - \paperheight }
    \draw_path_moveto:n { 5pt , 0pt }
    \draw_path_lineto:n { 20pt , 0pt }
    \draw_path_moveto:n { 0 , -5pt }
    \draw_path_lineto:n { 0 , -20pt }
    \draw_path_use:n { stroke }
  \draw_end:
}

\bool_if:NTF \l_zchinr_prepress_bool {
  \color_set:nnn { zchinrblue } { cmyk } { 1 , .5 , 0 , .5 }
} {
  \color_set:nnn { zchinrblue } { rgb } { 0 , .25 , .5 }
}

\bool_if:NT \l_zchinr_prepress_bool {
  \geometry{
    layout=a4paper,
    paperwidth=230mm,
    paperheight=317mm,
    layouthoffset=10mm,
    layoutvoffset=10mm
  }

  \AddToHook{shipout/foreground}{
    \__zchinr_cropmark_nw:
    \__zchinr_cropmark_ne:
    \__zchinr_cropmark_sw:
    \__zchinr_cropmark_se:

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
    \str_if_empty:NTF \l_zchinr_color_profile_path_str {
      \msg_warning:nn { zchinr } { no-color-profile } 
    } {
      \special{pdf:fstream ~ @cmykdata ~ (\l_zchinr_color_profile_path_str) ~ <<
        /N ~ 4 
        /Alternate/DeviceCMYK
      >>}
    }
    \str_if_empty:NTF \l_zchinr_color_profile_path_str {
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
          /OutputConditionIdentifier ~ (\l_zchinr_color_profile_name_tl)
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

\setlength{\LTpre}{0pt}
\setlength{\LTpost}{0pt}

% Style figure and table captions.
\RequirePackage[
  labelfont=bf,
  labelsep=period,
]{caption}

% Set up hyperreferences
\PassOptionsToPackage{hyphens}{url}
\bool_if:NTF \l_zchinr_prepress_bool {
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

\newpagestyle{zchinrempty}{
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

\titlecontents{part}[0pt]{\bigskip\Large\scshape\bfseries}{}{}{}[\bigskip]
\titlecontents{chapter}[15pt]{\vskip 5pt}{}{}{\hfill\contentspage}

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

% Style footnotes.
\tl_set:Nn \l__zchinr_fnmark_tl {
  \@thefnmark
}
\cs_set:Npn \@makefntext #1 {
  \parbox[t]{2em}{ \l__zchinr_fnmark_tl }
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

\NewDocumentCommand { \currentissueyear } { } {
  \tl_use:N \l_zchinr_issue_year_tl
}

\NewDocumentCommand { \currentissuevolume } { } {
  \tl_use:N \l_zchinr_issue_volume_tl
}

\NewDocumentCommand { \currentissuenumber } { } {
  \tl_use:N \l_zchinr_issue_number_tl
}

\NewDocumentCommand { \currentissuefirstpage } { } {
  \tl_use:N \g_zchinr_issue_first_page_tl
}

\NewDocumentCommand { \currentissuelastpage } { } {
  \tl_use:N \g_zchinr_issue_last_page_tl
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
  \bigskip
}

\NewDocumentCommand { \addressessection } { m o } {
  \phantomsection
  \addcontentsline{toc}{chapter}{\IfNoValueTF{#2}{#1}{#2}}
  \noindent{\Large #1}\par
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
  \tl_set:Nn \arraystretch { 1.75 }
  \noindent\begin{longtable}{
    @{} 
      >{\raggedright} 
      p{32.5mm}
    @{\hskip\columnsep} 
      >{\raggedright} 
      p{\dimexpr\textwidth - 32.5mm - 37.5mm - 1.5\columnsep}
    @{\hskip .5\columnsep} 
      >{\raggedleft\let\newline\\\arraybackslash} 
      p{37.5mm}
    @{}
  }
} {
  \end{longtable}
}

\NewExpandableDocumentCommand { \imprintspan } { m } {
  \multicolumn{2}{@{}p{\dimexpr\textwidth - 32.5mm - \columnsep}@{}}{#1}
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

\AtBeginDocument{
  \tl_gset:NV \g_zchinr_issue_first_page_tl \g__zchinr_issue_first_page_aux_tl
  \tl_gset:NV \g_zchinr_issue_last_page_tl \g__zchinr_issue_last_page_aux_tl
}

\AddToHook{shipout/before}{
  \str_if_eq:VnTF \l_zchinr_article_pagination_str { arabic } {
    \bool_if:NF \g__zchinr_issue_first_page_set_bool {
      \tl_gset:Ne \g_zchinr_issue_first_page_tl { \arabic{page} }
      \bool_gset_true:N \g__zchinr_issue_first_page_set_bool
    }
    \tl_gset:Ne \g_zchinr_issue_last_page_tl { \arabic{page} }
  }
}

\AtEndDocument{
  \iow_now:cx { @auxout } {
    \token_to_str:N \ExplSyntaxOn 
    \iow_newline:
    \tl_gset:Nn \token_to_str:N \g__zchinr_issue_first_page_aux_tl { \tl_use:N \g_zchinr_issue_first_page_tl } 
    \iow_newline:
    \tl_gset:Nn \token_to_str:N \g__zchinr_issue_last_page_aux_tl { \tl_use:N \g_zchinr_issue_last_page_tl } 
    \iow_newline:
    \token_to_str:N \ExplSyntaxOff
  }
}

\NewDocumentCommand { \printtitle } { } {
  \tl_if_empty:NF \l_zchinr_article_first_page_tl {
    \setcounter{page}{ \l_zchinr_article_first_page_tl }
  }
  \setcounter{footnote}{0}
  \startcontents[zchinrsub]

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

  \str_case_e:nnF { \l_zchinr_article_header_style_str } {
    { category } {
      \tl_set_eq:NN \l__zchinr_header_mark_even_tl \l_zchinr_article_category_tl
      \tl_set_eq:NN \l__zchinr_header_mark_odd_tl \l_zchinr_article_category_tl      }
    { title } {
      \tl_set_eq:NN \l__zchinr_header_mark_even_tl \l_zchinr_article_title_short_tl
      \tl_set_eq:NN \l__zchinr_header_mark_odd_tl \l_zchinr_article_title_short_tl
    }
    { author ~ title } {
      \tl_set:Nn \l__zchinr_header_mark_even_tl {
        { \itshape \l_zchinr_article_author_short_tl , } ~ \l_zchinr_article_title_short_tl
      }
      \tl_set:Nn \l__zchinr_header_mark_odd_tl {
        { \itshape \l_zchinr_article_author_short_tl , } ~ \l_zchinr_article_title_short_tl
      }
    }
  } {
    \msg_warning:nn { zchinr } { header-style-unknown } 
  }

  \tl_if_empty:NF \l_zchinr_article_doi_tl {
    \thispagestyle{zchinrtitle}
  }
  
  \tl_clear:N \l__zchinr_toc_entry_tl
  \str_case_e:nnF { \l_zchinr_article_toc_style_str } {
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

  \tl_set:Nn \thefootnote { \fnsymbol{footnote} }
  \tl_set:Nn \l__zchinr_fnmark_tl {
    \textsuperscript{\@thefnmark}
  }
  \__zchinr_article_title_block:
  \tl_if_empty:NF \l_zchinr_article_author_note_tl {
    \footnotetext{\l_zchinr_article_author_note_tl}
    \addtocounter{footnote}{-1}
  }
  \tl_set:Nn \thefootnote { \arabic{footnote} }
  \tl_set:Nn \l__zchinr_fnmark_tl {
    \@thefnmark
  }
  \@afterindentfalse\@afterheading
}

\cs_new:Npn \__zchinr_article_author: {
  \l_zchinr_article_author_tl
  \tl_if_empty:NF \l_zchinr_article_author_note_tl {
    \, \footnotemark
  }
}

\cs_new:Npn \__zchinr_article_category: {
  \bool_if:NF \l_zchinr_standalone_bool {
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

\cs_new:Npn \__zchinr_article_title_large: {
  \bool_if:NF \l_zchinr_standalone_bool {
    \phantomsection
    \addcontentsline{toc}{chapter}{
      \texorpdfstring{\l__zchinr_toc_entry_tl}{\l_zchinr_article_title_tl}
    }
  }
  \noindent\parbox[t][\l_zchinr_title_distance_dim][t]{\textwidth}{
    \raggedright{\huge\l_zchinr_article_title_tl}
    \clist_if_in:NnT \l_zchinr_article_title_style_clist { author } {
      \par\bigskip
      {\fontsize{11}{13}\selectfont\itshape \__zchinr_article_author: }
    }
  }
  \par\vskip 4mm\bigskip
}

\cs_new:Npn \__zchinr_article_title_small: {
  \bool_if:NF \l_zchinr_standalone_bool {
    \phantomsection
    \addcontentsline{toc}{chapter}{
      \texorpdfstring{\l__zchinr_toc_entry_tl}{\l_zchinr_article_title_tl}
    }
  }
  \noindent{\fontsize{11}{13}\selectfont\bfseries\l_zchinr_article_title_tl}
  \par\bigskip\nopagebreak
  \noindent{\itshape \__zchinr_article_author: }
  \par\bigskip\nopagebreak
}

\cs_new:Npn \__zchinr_article_abstract: {
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

\cs_new:Npn \__zchinr_article_subcontents: {
  \begin{subcontents}
    \printcontents[zchinrsub]{}{1}{}
  \end{subcontents}
  \par\vskip 12mm
}

\cs_new:Npn \__zchinr_article_title: {
  \clist_if_in:NnT \l_zchinr_article_title_style_clist { category } {
    \__zchinr_article_category:
  }
  \bool_if:NTF \l_zchinr_article_small_title_bool {
    \clist_if_in:NnT \l_zchinr_article_title_style_clist { title } {
      \__zchinr_article_title_small:
    }   
  } {
    \clist_if_in:NnT \l_zchinr_article_title_style_clist { title } {
      \__zchinr_article_title_large:
    }    
    \clist_if_in:NnT \l_zchinr_article_title_style_clist { abstract } {
      \__zchinr_article_abstract:
    }
    \clist_if_in:NnT \l_zchinr_article_title_style_clist { contents } {
      \__zchinr_article_subcontents:
    }
  } 
}

\cs_new:Npn \__zchinr_article_title_block: {
  \clist_if_in:NnF \l_zchinr_article_title_style_clist { title } {
    \clist_if_in:NnT \l_zchinr_article_title_style_clist { phantom } {
      \phantomsection
      \addcontentsline{toc}{chapter}{
        \texorpdfstring{\l__zchinr_toc_entry_tl}{\l_zchinr_article_title_tl}
      }
    }
  }
  \bool_if:NTF \l_zchinr_article_twocolumn_bool {
    \twocolumn[
      \__zchinr_article_title:
    ]
    \sloppy\flushbottom
    \tl_if_empty:NF \l_zchinr_article_doi_tl {
      \enlargethispage{-10mm}
      \afterpage{\enlargethispage{-10mm}}
    }
  } {
    \onecolumn
    \__zchinr_article_title:
    \tl_if_empty:NF \l_zchinr_article_doi_tl {
      \enlargethispage{-10mm}
    }
  }
}

\NewDocumentCommand { \inputpart } { s m } {
  \IfBooleanTF{#1}{
    \bool_if:NF \l_zchnr_standalone_bool {
      \input{#2}
    }
  }{
    \input{#2}
  }
}

\NewDocumentEnvironment { drawabsolute } { m +b } {
  \AddToHookNext{shipout/background}{
    \nullfont
    \draw_begin:
      \bool_set_false:N \l_draw_bb_update_bool
      \bool_if:NT \l_zchinr_prepress_bool {
        \draw_transform_shift:n { #1 }
      }
      #2
    \draw_end:
  }
  \thispagestyle{zchinrempty}
  \null
  \clearpage
} { }

\NewDocumentCommand { \drawrectangle } { o m m } { 
  \draw_scope_begin:
    \color_select:n { #1 }
    \draw_path_rectangle:nn { #2 } { #3 }
    \draw_path_use:n { fill }
  \draw_scope_end:
}

\NewDocumentCommand { \drawcoffin } { o m m O{hc} O{vc} +m } { 
  \draw_scope_begin:
    \color_select:n { #1 }
    \vcoffin_set:Nnn \l_tmpa_coffin { #3 } { #6 }
    \draw_transform_shift:n { #2 }
    \draw_coffin_use:Nnn \l_tmpa_coffin { #4 } { #5 }
  \draw_scope_end:
}

\NewDocumentCommand { \drawinput } { o m m } { 
  \draw_scope_begin:
    \color_select:n { #1 }
    \draw_transform_shift:n { #2 }
    \input{#3}
  \draw_scope_end:
}

\NewDocumentCommand { \oasymbol } { } { 
  \draw_begin:
    \draw_baseline:n { -.175em }
    \draw_scope_begin:
      \draw_evenodd_rule:
      \draw_path_circle:nn { 0 , 0 } { 1 }
      \draw_path_use_clear:n { fill }
      \draw_path_circle:nn { 0 , 0 } { 3 }
      \draw_path_circle:nn { 0 , 0 } { 2 }
      \draw_path_use_clear:n { fill }
      \draw_path_moveto:n { 2.75 , 0 } 
      \draw_path_lineto:n { 2.75 , 3.5 }
      \draw_path_arc:nnn { 0 } { 180 } { 2.75 }
      \draw_path_lineto:n { -1.75 , 3.5 }
      \draw_path_arc:nnn { 180 } { 0 } { 1.75 }
      \draw_path_lineto:n { 1.75 , 2 }
      \draw_path_close:
      \draw_path_use_clear:n { fill }
    \draw_scope_end:
  \draw_end:
}

% EOF