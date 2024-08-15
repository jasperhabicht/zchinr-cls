% === BOF

% ==========
% Document class for ZChinR papers.
% ==========
% This document class provides basic layout options for articles printed in ZChinR. 

% ==========

\def\@clsfiledate{2024/08/13}
\def\@clsfileversion{20240813v01}

\NeedsTeXFormat{LaTeX2e}
\ProvidesClass{zchinr}[\@clsfiledate\space (\@clsfileversion) by Jasper Habicht]

\RequirePackage{iftex}
\RequireXeTeX

\ExplSyntaxOn
  \NewDocumentCommand{\IfStrEqualTF}{ m m m m }{
    \str_if_eq:nnTF {#1} {#2} {#3} {#4}
  }
  \NewDocumentCommand{\IfStrEmptyTF}{ m m m }{
    \str_if_empty:nTF {#1} {#2} {#3}
  }
  \NewDocumentCommand{\IfStrEqualAnyTF}{ m m m m }{
    \clist_set:Nn \l_tmpa_clist {#2}
    \clist_if_in:NnTF \l_tmpa_clist {#1} {#3} {#4}
  }
\ExplSyntaxOff

\newif\if@zchinrprepress
\newif\if@zchinrtitlepage
\newif\if@zchinrstandalone
\newif\if@zchinrpagination

\DeclareOption{prepress}{\@zchinrprepresstrue}
\DeclareOption{titlepage}{\@zchinrtitlepagetrue}
\DeclareOption{standalone}{\@zchinrstandalonetrue}
\ProcessOptions\relax

% Load documentclass article.
\LoadClass[10pt,a4paper,twoside,twocolumn]{article}

% Set page margins. Set space between columns
\RequirePackage[
  inner=25mm,
  outer=20mm,
  top=30mm,
  bottom=20mm,
  headsep=5mm,
  footskip=5mm
]{geometry}
\setlength{\columnsep}{7.5mm}

\def\@zchinrcolorprofileiden{Custom}
\def\@zchinrcolorprofileinfo{none}
\def\@zchinrcolorprofilepath{}
\NewDocumentCommand{\colorprofile}{ m m m }{
  \def\@zchinrcolorprofileiden{#1}
  \def\@zchinrcolorprofileinfo{#2}
  \def\@zchinrcolorprofilepath{#3}
}

\if@zchinrprepress
  \RequirePackage[width=225truemm,height=312truemm,center,noinfo]{crop}
  \newcommand*\@zchinrcropmarkNW{
    \begin{picture}(0,0)
      \thinlines\unitlength1pt\linethickness{0.25pt}
      \put(-6,0){\line(-1,0){20}}
      \put(0,6){\line(0,1){20}}
    \end{picture}
  }
  \newcommand*\@zchinrcropmarkNE{
    \begin{picture}(0,0)
      \thinlines\unitlength1pt\linethickness{0.25pt}
      \put(6,0){\line(1,0){20}}
      \put(0,6){\line(0,1){20}}
    \end{picture}
  }
  \newcommand*\@zchinrcropmarkSW{
    \begin{picture}(0,0)
      \thinlines\unitlength1pt\linethickness{0.25pt}
      \put(-6,0){\line(-1,0){20}}
      \put(0,-6){\line(0,-1){20}}
    \end{picture}
  }
  \newcommand*\@zchinrcropmarkSE{
    \begin{picture}(0,0)
      \thinlines\unitlength1pt\linethickness{0.25pt}
      \put(6,0){\line(1,0){20}}
      \put(0,-6){\line(0,-1){20}}
    \end{picture}
  }
  \cropdef[]\@zchinrcropmarkNW\@zchinrcropmarkNE\@zchinrcropmarkSW\@zchinrcropmarkSE{zchinrcropmarks}
  \crop[zchinrcropmarks]

  \AtBeginDocument{
    \special{pdf:minorversion 3}
    \special{pdf:docinfo <<
      /GTS_PDFXVersion (PDF/X-1:2001)
      /GTS_PDFXConformance (PDF/X-1a:2001)
    >>}%
    \IfStrEqualTF{\@zchinrcolorprofilepath}{}{%
      \ClassWarning{zchinr}{No color profile defined.}%
    }{%
      \special{pdf:fstream @cmykdata (\@zchinrcolorprofilepath) <<
        /N 4^^J/Alternate/DeviceCMYK
      >>}%
    }%
    \IfStrEqualTF{}\@zchinrcolorprofilepath}{}{%
      \special{pdf:put @catalog <<
        /PageMode /UseNone
        /OutputIntents [<<
          /Type /OutputIntent
          /S /GTS_PDFX
          /OutputConditionIdentifier (Custom)
          /Info (none)
          /RegistryName (http://www.color.org/)
        >>]
      >>}%
    }{%
      \special{pdf:put @catalog <<
        /PageMode /UseNone
        /OutputIntents [<<
          /Type /OutputIntent
          /S /GTS_PDFX
          /DestOutputProfile @cmykdata
          /OutputConditionIdentifier (\@zchinrcolorprofileiden)
          /Info(\@zchinrcolorprofileinfo)
          /RegistryName (http://www.color.org/)
        >>]
      >>}%
    }%
    \special{pdf:put @thispage <<
      /MediaBox [0.00000 0.00000 637.79528 884.40945]
      /CropBox [0.00000 0.00000 637.79528 884.40945]
      /BleedBox [12.75591 12.75591 625.03937 871.65354]
      /TrimBox [21.25984 21.25984 616.53543 863.14961]
    >>}%
    \usepackage{atbegshi}
    \AtBeginShipout{%
      \special{pdf:put @thispage <<
        /MediaBox [0.00000 0.00000 637.79528 884.40945]
        /CropBox [0.00000 0.00000 637.79528 884.40945]
        /BleedBox [12.75591 12.75591 625.03937 871.65354]
        /TrimBox [21.25984 21.25984 616.53543 863.14961]
      >>}
    }
  }
\fi

% Set input encoding to UTF-8. Allows for direct input of umlauts in tex file. 
% Set main language to German but provide languange support for English as well, 
% allowing for setting the language within the tex file with \selectlanguage{spanish}.

% Following to be used with XeLaTeX
\RequirePackage{polyglossia} 
\setdefaultlanguage{english}
\setmainlanguage[spelling=new]{german}

\setmainfont{TeX Gyre Pagella}
\setsansfont{TeX Gyre Heros Cn}

\RequirePackage{xeCJK} 
\XeTeXlinebreaklocale "zh" 
\XeTeXlinebreakskip = 0pt plus 1pt minus .5pt
\setCJKmainfont{Noto Serif CJK SC}
\setCJKsansfont{Noto Sans CJK SC}

% Provide options for simplified and traditional Chinese. 
% Input with \zhs{simplified} or \zht{traditional}.
\NewDocumentCommand{\zhs}{ m }{\textup{#1}}
\NewDocumentCommand{\zht}{ m }{\textup{#1}}

% Unset double spacing after full stops. 
% Prevent single lines at the end or at the beginning of paragraphs.
\frenchspacing
\RequirePackage{indentfirst}
\RequirePackage{balance}
\RequirePackage{textpos}
\clubpenalty=10000
\widowpenalty=10000
\doublehyphendemerits=5000000

\setlength{\parindent}{10pt} % Default is 10pt
\setlength{\parskip}{2pt plus 1pt} % Default is 0pt plus 1pt

\RequirePackage{calc}
\RequirePackage{graphicx}
\RequirePackage[svgnames]{xcolor}
\if@zchinrprepress
  \selectcolormodel{cmyk}
\else
  \selectcolormodel{rgb}
\fi

\RequirePackage{array}
\RequirePackage{longtable}
\RequirePackage{booktabs}
\RequirePackage{enumitem}

\RequirePackage{tikz}
\usetikzlibrary{positioning,calc}

% Define commands. Commands set variables to the given value which will be used in other macros later.
% Example: \issue{2000}{1} sets the issue to 1 of year 2000.
% Example: \pages{2}{3} sets the first page of the article to 2, type will be set raggedright.
\def\@zchinrissueyear{}
\def\@zchinrissuenumber{}
\def\@zchinrdoi{}
\def\@zchinrfirstpage{1}
\def\@zchinrlastpage{1}
\def\@zchinrarticleauthor{}
\def\@zchinrarticleauthorshort{}
\def\@zchinrarticletitlepdf{}
\def\@zchinrarticletitlede{}
\def\@zchinrarticletitleshort{}
\def\@zchinrarticletitleen{}

\NewDocumentCommand{\issue}{ m m }{
  \def\@zchinrissueyear{#1}
  \def\@zchinrissuenumber{#2}
}
\NewDocumentCommand{\issueyear}{ m }{
  \def\@zchinrissueyear{#1}
}
\NewDocumentCommand{\issuenumber}{ m }{
  \def\@zchinrissuenumber{#1}
}
\NewDocumentCommand{\doi}{ m }{
  \def\@zchinrdoi{#1}
}
\NewDocumentCommand{\pages}{ s m m }{
  \IfBooleanT{#1}{\@zchinrpaginationtrue}{\@zchinrpaginationfalse}
  \def\@zchinrfirstpage{#2}\setcounter{page}{#2}
  \def\@zchinrlastpage{#3}
}
\NewDocumentCommand{\firstpage}{ s m }{
  \IfBooleanT{#1}{\@zchinrpaginationtrue}{\@zchinrpaginationfalse}
  \def\@zchinrfirstpage{#2}\setcounter{page}{#2}
}
\NewDocumentCommand{\lastpage}{ m }{
  \def\@zchinrlastpage{#1}
}
\if@zchinrtitlepage
  \newfontfamily\sansloosefamily[Mapping=tex-text,LetterSpace=2]{TeX Gyre Heros}
\else
  \RenewDocumentCommand{\author}{ m m }{
    \def\@zchinrarticleauthor{#1}
    \def\@zchinrarticleauthorshort{#2}
  }
  \RenewDocumentCommand{\title}{ o m m m }{
    \IfNoValueTF{#1}{
      \def\@zchinrarticletitlepdf{#2}
    }{
      \def\@zchinrarticletitlepdf{#1}
    }
    \def\@zchinrarticletitlede{#2}
    \def\@zchinrarticletitleshort{#3}
    \def\@zchinrarticletitleen{#4}
  }
\fi

\def\@zchinrarticleabstractde{}
\def\@zchinrarticleabstracten{}
\NewDocumentCommand{\articleabstract}{ +m +m }{
  \def\@zchinrarticleabstractde{#1}
  \def\@zchinrarticleabstracten{#2}
}

% Style header and footer. Header stays empty while footer will contain journal information and page number. 
\RequirePackage[pagestyles]{titlesec}
\def\@zchinrheadcentereven{}
\def\@zchinrheadcenterodd{}
\def\@zchinrthepage{\thepage}
\newpagestyle{main}{
  \sethead
    [ZChinR \@zchinrissueyear][\@zchinrheadcentereven][\@zchinrthepage]
    {\@zchinrthepage}{\@zchinrheadcenterodd}{ZChinR \@zchinrissueyear}
}
\pagestyle{main}

\NewDocumentCommand{\@zchinrsetfoothead}{ m }{
  \def\@zchinrheadcentereven{#1}
  \def\@zchinrheadcenterodd{#1}
  \def\@zchinrthepage{\thepage}
  \IfStrEqualTF{#1}{Aufsätze}{
    \def\@zchinrheadcentereven{\@zchinrarticleauthorshort}
    \def\@zchinrheadcenterodd{\@zchinrarticletitleshort}
  }{}
  \IfStrEqualTF{#1}{Kurze Beiträge}{
    \def\@zchinrheadcentereven{\@zchinrarticleauthorshort}
    \def\@zchinrheadcenterodd{\@zchinrarticletitleshort}
  }{}
  \IfStrEqualTF{#1}{Dokumentationen}{
    \def\@zchinrheadcentereven{\@zchinrarticletitleshort}
    \def\@zchinrheadcenterodd{\@zchinrarticletitleshort}
  }{}
  \IfStrEqualTF{#1}{Impressum}{
    \def\@zchinrthepage{}
  }{}
} 

\NewDocumentCommand{\articletype}{ m }{
  \IfStrEqualTF{#1}{Bibliography}{
    \def\@zchinrarticletypeprint{Dokumentationen}\def\@zchinrarticletype{#1}\@zchinrsetfoothead{Dokumentationen}
  }{
    \def\@zchinrarticletypeprint{#1}\def\@zchinrarticletype{#1}\@zchinrsetfoothead{#1}
  }
}

% Style section titles. 
\titleformat{\section}{\bfseries\fontsize{11}{13}\selectfont}{}{0pt}{}
\titleformat{\subsection}{\bfseries\fontsize{11}{13}\selectfont}{}{0pt}{}
\titleformat{\subsubsection}{\bfseries\fontsize{11}{13}\selectfont}{}{0pt}{}
\titlespacing*{\section}{0pt}{5pt plus 1pt}{5pt plus 1pt}
\titlespacing*{\subsection}{0pt}{5pt plus 1pt}{5pt plus 1pt}
\titlespacing*{\subsubsection}{0pt}{5pt plus 1pt}{5pt plus 1pt}

\titleformat{\paragraph}[runin]{\bfseries}{}{2em}{\noindent}
\titlespacing*{\paragraph}{0pt}{5pt plus 1pt}{1em plus 5pt minus 10pt}
\titleformat{\subparagraph}[runin]{\bfseries}{}{2em}{\noindent}
\titlespacing*{\subparagraph}{0pt}{5pt plus 1pt}{1em plus 5pt minus 10pt}

% Style footnotes.
\RenewDocumentCommand\@makefntext{ +m }{\leftskip=0em\hskip0em\@makefnmark\hskip1em#1}

% Style figure and table captions.
\RequirePackage[%
  labelfont=bf,%
  labelsep=period,%
]{caption}

% Create listing environment for references. 
\newlist{references}{enumerate}{1}
\setlist[references]{label*=[\arabic*]}

% Create command for author information box. 
\NewDocumentCommand{\authorbox}{ +m }{\noindent\fbox{\parbox{\columnwidth}{\small #1}}}

% Create command for automatically linked e-mail addresses. 
\NewDocumentCommand{\email}{ m }{\href{mailto:#1}{\protect\nolinkurl{#1}}}

% Create documentation environment. 
\NewDocumentEnvironment{documentation}{ }{
  \RenewDocumentCommand{\arraystretch}{1.25}
  \noindent\begin{longtable}{@{}>{\hskip 2em}{p}{54mm}@{\hskip 7mm}>{\hskip 15pt}{p}{115mm}@{}}
}{
  \end{longtable}
}

\NewDocumentCommand{\tablehead}{ O{c} O{} m o o }{%
  \IfNoValueTF{#4}{}{%
    \phantomsection%
    \addcontentsline{toc}{#4}{\IfNoValueTF{#5}{#3}{#5}}%
  }%
  \vspace{-\baselineskip}\begin{center}
    \IfStrEqualTF{#1}{l}{\raggedright}{}%
    \IfStrEqualTF{#1}{r}{\raggedleft}{}%
    \IfStrEqualTF{#2}{n}{}{\bfseries} #3
  \end{center}
}

% Create addresses environment.
\NewDocumentEnvironment{addresses}{ }{
  \setlength{\arrayrulewidth}{1pt}
  \RenewDocumentCommand{\arraystretch}{1.75}
  \noindent\begin{longtable}{@{}p{80.75mm}@{\hskip 7mm}p{88.25mm}@{}}
}{
  \end{longtable}\par
}

% Create imprint environment.
\NewDocumentEnvironment{imprint}{ }{
  \RenewDocumentCommand{\arraystretch}{1.25}
  \noindent\begin{longtable}{@{}p{30mm}@{\hskip 7mm}p{100mm}@{\hskip 4mm}>{\raggedleft\arraybackslash}p{35mm}@{}}
}{
  \end{longtable}\par
}

% Create commands for TOC. 
\NewDocumentCommand{\tocsection}{ m }{\bigskip\noindent{\Large\scshape\bfseries #1}}

\newlist{toc}{itemize}{1}
\setlist[toc]{leftmargin=15mm,label*={}}

% Override default abstract environment. 
\RenewDocumentEnvironment{abstract}{ }{%
  \begin{minipage}{\textwidth}\setlength{\parindent}{10pt}\noindent%
}{%
  \end{minipage}%
}

% Create command to typeset article title and abstract. 
\newlength{\headerdistance}
\setlength{\headerdistance}{35mm} % May be reset within single article document using \setlength.
\NewDocumentCommand{\@zchinrarticletitleabstract}{ }{%
  \if@zchinrstandalone\else%
    \phantomsection%
    \addcontentsline{toc}{chapter}{\@zchinrarticletitlede}%
  \fi%
  \IfStrEqualTF{\@zchinrarticletype}{Dokumentationen}{%
    \noindent\parbox[t][\headerdistance][t]{\textwidth}{\raggedright\huge\@zchinrarticletitlede}\par\vskip 4mm\bigskip%
  }{%
    \IfStrEqualTF{\@zchinrarticletype}{Bibliography}{%
      \noindent\parbox[t][\headerdistance][t]{\textwidth}{\raggedright\huge\@zchinrarticletitlede\par\bigskip
      \fontsize{11}{13}\selectfont\itshape\@zchinrarticleauthor}\par\vskip 4mm\bigskip
    }{%
      \noindent\parbox[t][\headerdistance][t]{\textwidth}{\raggedright\huge\@zchinrarticletitlede\par\bigskip
      \fontsize{11}{13}\selectfont\itshape\@zchinrarticleauthor}\par\vskip 4mm\bigskip
      \IfStrEmptyTF{\@zchinrarticleabstractde}{}{% 
        \noindent{\itshape\bfseries\fontsize{11}{13}\selectfont Abstract}\par\vskip 8pt%
        \noindent\begin{abstract}%
          \itshape\@zchinrarticleabstractde%
          \IfStrEmptyTF{\@zchinrarticleabstracten}{}{%
            \par\bigskip%
            \noindent\textbf{\@zchinrarticletitleen} ---
            \@zchinrarticleabstracten%
          }%
        \end{abstract}\vskip 5mm%
      }
    }
  }
}

% Create command to balance twocol at end of article and shift footnotes. 
% Example: \balancehere[1] shifts footnotes upwards by 1 line.
\def\balancelineskip{0}
\NewDocumentCommand{\balancehere}{ O{0} }{
  \enlargethispage{-#1\baselineskip}
  \newpage
  \def\balancelineskip{#1}
}

% Create command to force justify. 
\NewDocumentCommand{\forcejustify}{ }{\linebreak\vspace*{-1\baselineskip}}
\NewDocumentCommand{\noindentchinese}{ }{\hskip -2em}
\NewDocumentCommand{\noindentgerman}{ }{\hskip -15pt}

% Create command to insert article title and abstract in a onecolumn area at top of first page. 
\NewDocumentCommand{\@zchinrmakearticletitleinner}{ m }{%
  \IfStrEqualTF{#1}{type+title}{%
    \if@zchinrstandalone\else%
      \phantomsection%
      \addcontentsline{toc}{part}{\@zchinrarticletypeprint}%
    \fi%
    \noindent\rule{\textwidth}{1pt}\par\vskip 2mm%
    \ifodd\@zchinrfirstpage
      \noindent\parbox[t]{\textwidth}{\raggedleft\huge\scshape\bfseries\@zchinrarticletypeprint}\par%
    \else
      \noindent\parbox[t]{\textwidth}{\raggedright\huge\scshape\bfseries\@zchinrarticletypeprint}\par%
    \fi
    \noindent\rule{\textwidth}{1pt}
    \par\vskip 7mm%
    \@zchinrarticletitleabstract
  }{
    \IfStrEqualTF{#1}{type}{%
      \if@zchinrstandalone\else%
        \phantomsection%
        \addcontentsline{toc}{part}{\@zchinrarticletypeprint}%
      \fi%
      \noindent\rule{\textwidth}{1pt}\par\vskip 2mm%
      \ifodd\@zchinrfirstpage
        \noindent\parbox[t]{\textwidth}{\raggedleft\huge\scshape\bfseries\@zchinrarticletypeprint}\par%
      \else
        \noindent\parbox[t]{\textwidth}{\raggedright\huge\scshape\bfseries\@zchinrarticletypeprint}\par%
      \fi
      \noindent\rule{\textwidth}{1pt}\par\vskip 7mm%
    }{
      \@zchinrarticletitleabstract
    }
  }
}

\NewDocumentCommand{\makearticletitle}{ O{title} }{%
  \IfStrEqualAnyTF{#1}{Dokumentationen,Inhalt,Adressen,Impressum}{ 
    \onecolumn
    \@zchinrmakearticletitleinner{#1}
  }{
    \twocolumn[
      \begin{@twocolumnfalse}
        \@zchinrmakearticletitleinner{#1}
      \end{@twocolumnfalse}
    ]
  }
}

% Load hyperref package which provides for clickable URLs and styling. Set URL font to main font.
\PassOptionsToPackage{hyphens}{url}
\if@zchinrprepress
  \RequirePackage[bookmarks=false]{hyperref}
  \RequirePackage{datetime}
  \AtBeginDocument{
    \NoHyper
    \hypersetup{% 
      pdftitle={ZChinR \@zchinrissueyear{} --- S.~\@zchinrfirstpage{}--\@zchinrlastpage},
      pdfinfo={
        ModDate={D:\pdfdate},
        Trapped={False},      
      }
    }
  }
\else
  \RequirePackage[hidelinks]{hyperref}
\fi
\RenewDocumentCommand{\toclevel@part}{ }{-1}
\urlstyle{same}
\DeclareUrlCommand\Hurl{\def\UrlLeft{<\,}\def\UrlRight{\,>}}
\g@addto@macro{\UrlBreaks}{\do\^\do\_}

% Kern double-slash in URLs. 
\Urlmuskip = 0mu
\newmuskip\@zchinrurlslashmuskip
\@zchinrurlslashmuskip = -3mu minus 1mu
\g@addto@macro\UrlSpecials{\do\/{\@zchinrurlkernslash}}
\def\@zchinrurlkernslash{\futurelet\@zchinrurlslashnext\@zchinrurlunspaceslash}
\def\@zchinrurlunspaceslash{%
  \mathchar8239 %
  \ifx\@zchinrurlslashnext/\mskip\@zchinrurlslashmuskip\fi%
}

% Check whether needed commands are set properly.
\AtBeginDocument{%
  \IfStrEmptyTF{\@zchinrissuenumber}{
    \ClassError{zchinr}{No \string\issue\space given}{Command \string\issue\space (or \string\issueyear\space and \string\issuenumber\space) has to be included in preamble}
    \stop
  }{}
  \IfStrEmptyTF{\@zchinrissueyear}{
    \ClassError{zchinr}{No \string\issue\space given}{Command \string\issue\space (or \string\issueyear\space and \string\issuenumber\space) has to be included in preamble}
    \stop
  }{}
}

\if@zchinrtitlepage\else
  % Check whether \lastpage is identical with last page of document.
  \AtEndDocument{%
    \IfStrEqualTF{\thepage}{\@zchinrlastpage}{}{
      \ClassWarning{zchinr}{Last page number of document (\thepage) not identical with last page stated in \string\pages\space (\@zchinrlastpage)}
    }
  }
\fi

% === EOF