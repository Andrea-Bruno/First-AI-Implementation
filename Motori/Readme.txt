- Readme File - Zurich 25th Apr. 1999




  ================
  - Cilian v4.13 -
  ================




  1. Copyright & License
  ----------------------

LIMITED WARRANTY

THE PROGRAM IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR
IMPLIED, INCLUDING, BUT NOT LIMITED TO WARRANTIES OF MERCHANTABILITY OR FITNESS
FOR A PARTICULAR PURPOSE.  IN NO EVENT WILL THE AUTHOR or AUTHORS BE LIABLE TO
YOU FOR ANY DAMAGES, INCLUDING INCIDENTAL OR CONSEQUENTIAL DAMAGES, ARISING OUT
OF THE USE OF THE PROGRAM, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGES.
YOU ACKNOWLEDGE THAT YOU HAVE READ THIS LICENSE, UNDERSTAND IT AND AGREE TO BE
BOUND BY ITS TERMS AS THE COMPLETE AND EXCLUSIVE STATEMENT OF THE AGREEMENT
BETWEEN US, SUPERSEDING ANY PROPOSAL OR PRIOR AGREEMENT, ORAL OR WRITTEN, AND
ANY OTHER COMMUNICATIONS BETWEEN US RELATING TO THE SUBJECT MATTER OF THIS
LICENSE.

                                 --*--

This version of Cilian is still a freeware, this means that you can duplicate it
and distribute it to other people but you must include all files in the original
distribution. Those files must be left unmodified, especially the executables
may not be changed by applying "patch" or "crak" files on them. Anyone
distributing this product for any kind of remuneration must first contact me for
authorization.




  2. What is it?
  --------------

Cilian is a chess program. Its purpose is to play chess and analyze games.

Cilian v4.09 plays with a rating of 2070 Elo on a K6-300Mhz (LCT II v1.21 (c)
Frédéric Louguet & La Puce Echiquéenne)

If you are programmer and want to use Cilian's chess engine in your application,
just email me! Cilian was written in GNU-C and it offers an easy fashion to
implement best move searching. Its function library was actually the subject of
a semester project at the ETH - Zurich.

The searcher IS NOT an alpha-beta searcher but is using MT boundary tests. (see
http://www.cs.vu.nl/~aske/) The driver of these tests is a special MT Driver
function who searches the best move with no need of searching its exact value.
Today the most chess engines use alpha-beta searchers like PVS-NegaScout, so I
wanted to be a little bit different and try to build an MT engine, since MT
appeared to be a new wave! I must say that MT posed me some hard problems (like
hashtables that need to be differently used by MT that by alpha-beta) and
solving bugs was a dreadful nightmare (since with MT the tree is scanned a lot
of times, you first need to know where and at which scan the error occurred,
before trying to understand it..). Today I finally got it :)! If you are
interested to discuss about MT implementation, you are welcome! Especially, I
would be happy to have a chat with somebody who is writing or wrote a MT based
program.

Cilian's features:

- Null-Window Tests (MT) based search engine,
- Quiescence with static 1-ply null-move and swap-off (static exchanges) (*),
- Transpositions Hashtables with Separated Lower and Upper Bounds Storage (*),
- Enhanced Transpositions Cut-offs,
- Futility Cut-offs at Frontier and Prefrontier nodes,
- Null-Move Heuristic (nonrecursive, R=2),
- Razoring (yet disabled),
- History Heuristic (useful but not optimal with MT),
- Static Extensions (normal (checks and promotions) or High Selective (according
  to the tactical value of the move)) (*),
- Anantharaman's Extensions (yet disabled),
- Swap-offs (static exchanges),
- Menace Detection (direct, uncover, second) (*),
- Material Exchange Tropism (*),
- Evaluation function (with 3-plies static exchanges, detection of secure
  squares, 2-plies static null-move, and many more...) (*),
- Evaluation Hashing (yet disabled),
- Lazy Evaluation,
- Draw Detection (repetition, 50 moves rule, no mating material),
- MT Driver with Range of Values (*),
- Iterative Deepening,
- Pondering,
- Tourney Time Control (*),
- Opennings Book.

(*): original ideas written exclusively for Cilian !!

To do: singular extensions (look at Fortress, the program of Alessandro Damiani, 
to see how strong they are!), re-tuning of the evaluation parameters according 
the root position, ...and the whole program structure  is not optimal (no 
bitboards, etc.) but I cannot bring myself to rewrite the whole thing (>10000 
lines!!)...




  3. Playing with WinBoard
  ------------------------

Cilian does not have a graphical user interface, so if you dislike using the
console mode, you should use WinBoard or another compatible interface.
(http://www.research.digital.com/SRC/personal/Tim_Mann/chess.html#XBoard)

WinBoard 4.0 specifications:

  -cp -fcp "Cilian.exe -xboard -hash18 -mvnw"
  -fd C:\cilian_dir -scp GNUChess
  -initString="new\nset 3 6\n"

In WinBoard, you cannot use all menu functions, but 'ponder next move', 'show
thinking' and 'retract move' are supported, you can do 'load/save/edit
position/game', too.




  4. Playing with the Console
  ---------------------------

Command line options of Cilian:

 -xboard                : put Cilian in X-/WinBoard or ICS mode.
 -hashX                 : use 2^n entries in two hashtables.
 -easy                  : Cilian is fixed in easy mode (no pondering).
 -hard                  : Cilian is fixed in hard mode (pondering).
 -mvnw                  : enable 'move now!' feature, recommended with
                          X-/WinBoard.
 -nobk                  : disable the openings book.
 -crbk                  : create the openings book.

Try to increase or decrease the hashtable size according to your amount of
memory. (-hash18=13MB, -hash19=21MB, -hash20=33MB and so forth)

The 'move now!' feature is always enable if Cilian is in hard mode, but not in 
easy mode (default); the purpose of this is that you can use a batch file (like 
"Cilian -easy < myscript.txt") and write in "myscript.txt" everything (like 
"load blabla.pgn\n go\n load next.pgn\n set 6 6\n go\n"). If 'move now!' were 
always enable, a batch file would never be usable, because with a batch file 
stdin is never empty and Cilian checks this to know if he has to stop searching. 
(see appendix for more informations on scripts)

The 'openings book' is a new feature of Cilian! For this you need to have the 
file "openings.dat" in the working directory of Cilian. If you want, you can 
edit and modify the file "openings.txt" and recompile it with the command 
"Cilian -crbk". This allow you to build your own opennings book!

                                 --*--

Open a console and launch Cilian, you should get the following:

  Cilian v#.##
  new game

And then use the following commands:
(the commands with a (*) are used by X-/WinBoard)

+---------------------------+
| le(vel) MOVE TOTAL INC (*)|
+---------------------------+
Set the time control mode:

MOVE = X  means that after X moves the clocks receive the amount of minutes
          specified in TOTAL again.
MOVE = 0  means that the whole game (no matter how many moves it takes) has to
          be played inside the amount of minutes specified in TOTAL.
TOTAL = X means that the clocks have X minutes at the beginning. (you can use
          5.5 if you want 5 minutes 30 seconds)
TOTAL = 0 means (if INC = 0) that there is no time control, and the search
          never stops because of time limits.
INC = X   means that after each played move the clocks receive an increment
          time of X seconds.

Note that there is no time check for the human side, so if you set Cilian to
play in 5 minutes, you are still free to play so slow as you like.

Examples:

  level 0 5 0   -> Cilian plays the whole game inside 5 minutes.
  level 40 5 0  -> Cilian plays 40 moves inside 5 minutes; 5 minutes are added
                   to its clock after the 40th, the 80th move and so on.
  level 0 2 12  -> Cilian plays the whole game inside 2 minutes and receives a
                   bonus of 12 seconds for each move it plays.
  level 0 0 0   -> Cilian does not use time limits (for analyze purposes).


+-----------+
| se(t) D Q |
+-----------+
Set the depth control:

D is the brute force depth, that is the number of half-moves (plies) that Cilian
has to investigate before it chooses a move. On this depth Cilian adds some
extensions if it is worth it, and finally, it adds Q plies of quiescence search
(except for D=1, which is the turbo mode). Usually, you do not need to change Q.

Note: if you want to make a search at a fixed level (for analyze purpose), you
should type 'level 0 0 0' to disable the time control, else Cilian try to save
time and does not always compute until D is reached. (see appendix)

+--------+
| hel(p) |
+--------+
Display the summary of commands and the current settings.

+-------+
| go (*)|
+-------+
Set the computer to play with the current color and begin the search of its
move. (use Ctrl-C to interrupt the search)

+-------------------------------+
| e5 (or another valid move) (*)|
+-------------------------------+
Make the move, after which the computer begins automatically its move search if
it is enable. The supported notations are short algebraic, standard algebraic
and Auto232:

  'e4', 'e2e4' or 'e2-e4' (Pawn),
  'Ne2', 'g1e2' or 'g1-g2' (Knight),
  'Nbc3', 'b1c3' or 'b1-c3' (Knight in b-column),
  'O-O', 'e1g1' or 'e1-g1' (King's side castle),
  'O-O-O', 'e1g1' or 'e1-g1' (Queen's side castle),
  'Nxd4', 'e2d4' or 'e2xd4' (capture)
  'a8=Q', 'a7a8Q', 'a7-a8/Q' or 'a7-a8=Q'(promote).

Note that the capitals can always be replaced by small letters.

+--------+
| st(op) |
+--------+
Disable the computer. (to interrupt a search, use Ctrl-C)

+--------+
| ev(al) |
+--------+
Display the board and some informations of the current position.
Note that the '@' are White pawns and the '*' are Black pawns.

Example:
----------------------------------------------------
     a   b   c   d   e   f   g   h
   +---+---+---+---+---+---+---+---+
 8 | r | . |   | . |   |.r | k | . | 8
   +---+---+---+---+---+---+---+---+
 7 |.* | * | . | n | . | * |.* |   | 7
   +---+---+---+---+---+---+---+---+
 6 |   | . | * | . | * | . |   |.* | 6
   +---+---+---+---+---+---+---+---+
 5 |.q |   | . | * | . |   | . |   | 5
   +---+---+---+---+---+---+---+---+
 4 |   | . |   |.@ |   |.@ |   | . | 4
   +---+---+---+---+---+---+---+---+
 3 |.@ |   |.b | B |.@ |   | . |   | 3
   +---+---+---+---+---+---+---+---+
 2 |   |.@ | @ |.Q |   | . | @ |.@ | 2
   +---+---+---+---+---+---+---+---+
 1 |.R |   | . |   | . | R |.K |   | 1
   +---+---+---+---+---+---+---+---+
     a   b   c   d   e   f   g   h
 13. ... b4c3  is the last move
White to move
display details? (1/0): 0
- recapture on square (3,3): +3000
- menace: level=1 / direct=8977(4,2) / uncover=0(0,0) / second=977(4,4)
- material balance (with exchange tropism): -3448
- White: position=3984 material=29000 / Black: position=4368 material=32000
- global evaluation for White: -832
----------------------------------------------------

+-----------+
| in(verse) |
+-----------+
Set the flip view mode. (if you like seeing the Black side at the bottom)

+----+
| pv |
+----+
Set the Principal Variation, which is displayed while searching (it is actually 
the same flag as in the 'pv' command).
Note that it is approximate and could be incorrect, but it is still a good way
to help analyzing a position.

+---------------+
| rem(ove) X (*)|
+---------------+
Go X plies backward in the game.

+------------+
| rep(lay) X |
+------------+
Go to the Xth move in the game.

+---+
| n |
+---+
Go to next ply.

+---+
| p |
+---+
Go to previous ply.

+-----------+
| un(do) (*)|
+-----------+
Same as the 'p' command.

+---------+
| fl(ush) |
+---------+
Clear all hashtables and history heuristic tables.

+--------+
| new (*)|
+--------+
Begin a new game. (without automatically clearing the hashtables and history
heuristic tables)

+--------+
| sa(ve) |
+--------+
Save the game into a .pgn file (readable by WinBoard for ex.).

Note: if the initial position was nonstandard (modified by 'edit') then the
resulted saved file is only compatible with Cilian (because X-/WinBoard uses
.fen files and Cilian cannot save in that format). And if you edit and save a
position within X-/WinBoard, you will be able to load it with X-/WinBoard but
not with Cilian in console mode.

Look at the examples belove. X-/WinBoard only compatible means that Cilian need to be
interfaced with X-/WinBoard in order to recognize the specified file.

Example of an X-/WinBoard and Cilian compatible .pgn file:
----------------------------------------------------
[Event "?"]
1. d4 b5
2. c3 Bc7
*
----------------------------------------------------

Example of a Cilian only compatible .pgn file:
----------------------------------------------------
[Event "Cilian game"]
{ Cilian edited position: $E
-4 -2 -3 -5 -6 -3 -2 -4  1
-1 -1 -1 -1 -1 -1 -1 -1  0
 0  0  0  0  0  0  0  0  3
 0  0  0  0  0  0  0  0  3
 0  0  0  0  0  0  0  0  7
 5  0  0  0  0  0  0  0  4
 1  1  1  1  1  1  1  1  0
 4  2  3  5  6  3  2  4  4
}
  1. e2e4  d7d5
  2. a3a4  *
----------------------------------------------------

Example of an X-/WinBoard only compatible .pgn file:
----------------------------------------------------
[Event "?"]
[Result "*"]
[FEN "rb1qkbnr/p2ppppp/1p6/8/8/8/PPPPPPPP/RNBQKBNR w KQkq - 0 1"]
[SetUp "1"]
{--------------
r b . q k b n r
p . . p p p p p
. p . . . . . .
. . . . . . . .
. . . . . . . .
. . . . . . . .
P P P P P P P P
R N B Q K B N R
white to play
--------------}
1. d4 b5 2. c3 Bc7
*
----------------------------------------------------

Example of an X-/WinBoard and Cilian compatible .fen file:
----------------------------------------------------
KQkq1kr5/p2p2p1/p7/1p2b3/3P3P/6P1/5Q2/6K1 b - - 0 1
----------------------------------------------------

+--------+
| lo(ad) |
+--------+
Read a .pgn game file or a .fen position file.
To load a position in .fen format, you must use the extension .fen, because
Cilian assumes that all files without the .fen extension are .pgn files.

+-----------+
| edi(t) (*)|
+-----------+
Let you change one or more rows of the chessboard. Just type the row number and
then the eight square values (0=free, 1=White pawn, 2=White Knight, 3=White
bishop, 4=White rook, 5=White Queen, 6=White king and similarly 1,-2,-3,-4,-5,-6
for the Black pieces). Finally, enter the castle possibilities (0=none, 1=King's
side, 2=Queen's side and 3=both side) for the two White and Black and then type
the color (1=White and -1=Black) of the player on move.
Note: a position is invalid if, for ex., the are more than 16 White pieces.

+---------+
| hard (*)|
+---------+
Enable pondering, this means that Cilian uses your think time, too.
Note that if you use the command line options -easy or -hard, then you cannot
change this setting anymore.

+---------+
| easy (*)|
+---------+
Disable pondering. (see above)

+-----------+
| au(to232) |
+-----------+
This was initially the Auto232 support but actually, this is something other
(because Auto232 support is not yet implemented): if you enable this mode all
input are read in the file 'hismove.txt' and all output are written in
'mymove.txt'. This let you command Cilian with another program. Hit Ctrl-C or
use the command auto232 a second time to exit this mode.

                                 --*--

The following commands are not for the user but for X-/WinBoard or ICS support:

+----------+
| wb(oard) |
+----------+
Set X-/WinBoard and ICS support. Actually, if you enable this setting, the main
change in the behavior of Cilian is that it becomes quieter.

+-------------+
| ti(me) X (*)|
+-------------+
Set the remaining time to X/100 seconds.

+-------------+
| ot(im) X (*)|
+-------------+
This command comes for compatibility but does nothing.

+----------+
| sw(itch) |
+----------+
Change the color of the computer side (used by RoboFICS).

+------------+
| wh(ite) (*)|
+------------+
Set the computer playing White.

+------------+
| bl(ack) (*)|
+------------+
Set the computer playing Black.

+-----------+
| po(st) (*)|
+-----------+
Enable the 'show thinking' feature for X-/WinBoard (it is actually the same 
flag as in the 'pv' command).
Note that the gived value is not the minimax value, but it is a boundary (the
minimax value is greater or equal to the boundary).

+-------------+
| no(post) (*)|
+-------------+
Disable the 'show thinking' feature.

+-----------+
| edi(t) (*)|
+-----------+
The 'edit' command in X-/WinBoard mode has the usage specified in 
"engine-intf.txt" in WinBoard's package.

+-------------+
| res(ult) (*)|
+-------------+
The 'result' command does the same as 'stop' with the only difference that the 
characters right to 'result' are dropped.

                                 --*--

Note: you can use some abbreviated forms, for ex.: 'le' instead of 'level', and
you can write sequences of commands on one line. For example:

  lo a.pgn rep 13 fl go

This sequence will open the file a.pgn, go to the 13th move and clear the tables,
then it will start the chess engine.




  5. Credits
  ----------

The author of Cilian is studying Computational Science and Engineering at the
Swiss Federal Institute of Technology of Zurich (ETHZ). Contact him at the
following address:

  François Scheurer
  Universitaetstr.21
  CH-8006 Zurich
          ------

  email: 1st: fscheure@g26.ethz.ch
         2nd: milin@access.ch 


See Djordje's webpage about Computer Chess! He has some strong freeware chess
programs that you can download, included Cilian. You can see the tourney
activity that he maintains between some 20 engines. (Cilian is at the bottom
with something like 2000 ELO, but he is proud to be there with so famous
engines!)

  http://www.cent.co.yu/chess/

The source code can be build for all platforms which have a GNU-C compiler or
any ANSI-C compiler. Call me if you are interested in Cilian for another
platform that the one joint in this package.

Thanks to Prof. Jürg Nievergelt and his assistant Christoph Wirth, who accepted
my request for doing a semester project in Computer Chess.
Special thanks to Alessandro Damiani for his advices in computer chess.




  6. Appendix
  -----------

'set' command:

For ex.: if you chose 'level 40 60 12', Cilian knows he has 60*60/40+12=102
sec available for each moves, but will not start a new iteration after having
used more than 25% of the time, so 25 sec here. He starts then the search with
iterating deepening, that is depth=1, 2, 3 and so on, and if he just completed
the iteration with depth=4 after having used 30 sec, he will not begin the next
iteration, saving 72 sec. But if you wrote 'set 5 6' then Cilian must begin the
next iteration (depth=5), risking to be interrupted after 102 sec without
having completed the iteration with depth=5. Note: if the time control is not
disabled, there are some other special rules than the "25%-rule" that can
prevent Cilian to reach the depth specified with 'set'.

                                 --*--

'winboard.ini' file:

If you like playing match between different chess engines, here is an example
for the bottom of "winboard.ini" file.

/firstChessProgramNames="\"Cilian -hard -hash17 -xboard -mvnw\" /initString=\"new\\nset 3 6\\n\"
\"Cilian -easy -hash18 -xboard -mvnw\" /initString=\"new\\nset 3 6\\neasy\\n\"
GNUChess /initString=\"new\\neasy\\n\"
fortress -fd d:\\devstudio\\myprojects\\cilian_project\\engines\\fortress
"
/secondChessProgramNames="GNUChess /initString=\"new\\neasy\\n\"
\"Cilian -easy -hash18 -xboard -mvnw\" /initString=\"new\\nset 3 6\\neasy\\n\"
ssechess.exe -sd d:\\devstudio\\myprojects\\cilian_project\\engines\\ssechess
\"tris4018 wb hash=4m\" -sd d:\\devstudio\\myprojects\\cilian_project\\engines\\tristram
wcrafty- -sd d:\\devstudio\\myprojects\\cilian_project\\engines\\crafty
\"tcb tcbxini.cui\" -sd d:\\devstudio\\myprojects\\cilian_project\\engines\\tcb
fortress -sd d:\\devstudio\\myprojects\\cilian_project\\engines\\fortress
"

                                 --*--

scripts:

You can use scripts with Cilian, with the condition that pondering is off and
the -mvnw option is not specified.

Example of a Cilian script file:
----------------------------------------------------
wboard
post
level 10 100 0
set 31 6

new
flush
load p1.fen
go

new
flush
load p2.fen
go

new
flush
load p3.fen
go

quit
----------------------------------------------------

This script will use 10 minutes for each position.
To use it, type:

  Cilian < myscript.txt > results.txt

NEVER forget the 'quit' command at the end, else Cilian will write a huge
'results.txt' file with the line "Error (unknown command):"!

                                 --*--




  7. History
  ----------

The beginning of the period of development was March 1997 until June 97, after
what I did not have the time to program further. In April 98 I could continue the
development until October 98. So now, after 10 months of hard sweat, I can guess
that Cilian is okay for its first release.

25th Oct. 1998: release v3.97.

27th Oct. 1998: release v4.00.

 1st Nov. 1998: v4.01;
                -support 'edit/load/save position' and 'post' in X-/WinBoard.

 2nd Nov. 1998: v4.02;
                -MT faster with null-move R=2 and negative static extension of
                 2/8 of ply for the nonmenacing moves.

 7th Nov. 1998: v4.03;
                -better static extensions;
                -'load' was unable to properly read the castle possibilities
                 with .fen files (bug).

11th Nov. 1998: v4.04;
                -high selective search (new) in MT, use brute force only up to
                 depth/2, use static depth extensions/reductions in the deeper
                 search according to the tactical values of the moves;
                -the draw by 3 repetitions was not always detected (bug);
                -the time control made sometimes to quick searches (bug).

17th Nov. 1998: v4.05;
                -the extensions/reductions system is modified (better but need
                 to be more improved);
                -the repetition detection (should) uses a better algorithm.

19th Nov. 1998: v4.06;
                -better evaluation for pawns finals (square rule implemented,
                 etc.) (new);
                -'edit/load position' did not verify if castle possibilities were
                 valid (bug).

25th Nov. 1998: v4.07;
                -better board display for console mode (thanks to Lienne Chung);
                -lazy evaluation is implemented (new);
                -some additions/modifications to the pawns structure evaluation;
                -the 'neither player has material to mate' detection was
                 incorrect programmed (bug);
                -the command line option '-hashXX' was unusable (ugly bug!);
                -the 'post' command displayed a wrong estimation (bug);
                -Anantharaman's extension (new).

17th Dec. 1998: v4.08;
                -use a new compiler (Cygnus mingw32 GNU-C compiler) so
                 you don't need the big .dll file anymore;
                -the search could be too much reduced (bug);
                -two or three little bugs fixed (one for endgame evaluation);
                -the repetition detection still fails with some cases (will be
                 fixed soon).

21th Dec. 1998: v4.09;
                -something was wrong with the time control in console mode
                 (bug);
                -Cilian could crash in some rare cases, for ex.: if the move
                 O-O-O was made by black while the white king was in d1, then
                 the menace detection didn't detect the check, afterwhich a
                 null-move could cause a king capture (!) and then the program
                 could crash or just continue (with wrong results), thanks to
                 Kai Skibbe for the report of this bigbadbug!;
                -a weight value in evaluation caused strange play, the old
                 value is restored from v4.07.
                -the repetition detection should be fine now (bug);
                -king safety and some other variables in evaluation improved;
                -king castle evaluation refined.

11th Jan. 1999: v4.10;
                -the time control got new incoherences with the last
                 modifications (bug);
                -this version will be the last until April... in the next
                 version, two new things will be implemented:a) a small,
                 customizable, openings book and b) better play in tactical
                 positions.

 8th Feb. 1999: v4.11;
                -the search could make precision errors because of lazy eval
                 with MT (bug);
                -Cilian stops playing anymore if a draw occurs (for ex. 3
                 repetitions), but this is a problem with ICS, because if Cilian
                 plays a move which causes a 3rd repetition, he sends to ICS his
                 move and then sends the "draw" string, after which ICS should
                 terminate the game... but sometimes (because of lag) ICS
                 received the opponent's move exactly before the "draw" string,
                 avoiding the Cilian's demand for a draw to be considered; for 
                 this reason Cilian will not stop playing anymore if X-/WinBoard 
                 mode is enabled;
                -command 'result' added (new).

 4th Mar. 1999: v4.12;
                -a variable could be unitialized in some cases (if quiescence 
                 was disable) and causes a little random factor in the search 
                 (bug);
                -the null-move heuristic is now disable until next version.

26th Apr. 1999: v4.13;
                -little fix in the extension search (bug);
                -little improvement in the extension search;
                -Opennings Book implemented (new).