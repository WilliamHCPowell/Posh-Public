#
# script to print text as pairs of letters, and fives
#

param ($Stuff)

cls

$PlainText = @"
1 In the beginning was the Word, and the Word was with God, and the Word was God.

2 The same was in the beginning with God.

3 All things were made by him; and without him was not any thing made that was made.

4 In him was life; and the life was the light of men.

5 And the light shineth in darkness; and the darkness comprehended it not.

6 There was a man sent from God, whose name was John.

7 The same came for a witness, to bear witness of the Light, that all men through him might believe.

8 He was not that Light, but was sent to bear witness of that Light.

9 That was the true Light, which lighteth every man that cometh into the world.

10 He was in the world, and the world was made by him, and the world knew him not.

11 He came unto his own, and his own received him not.

12 But as many as received him, to them gave he power to become the sons of God, even to them that believe on his name:

13 Which were born, not of blood, nor of the will of the flesh, nor of the will of man, but of God.

14 And the Word was made flesh, and dwelt among us, (and we beheld his glory, the glory as of the only begotten of the Father,) full of grace and truth.

15 John bare witness of him, and cried, saying, This was he of whom I spake, He that cometh after me is preferred before me: for he was before me.

16 And of his fulness have all we received, and grace for grace.

17 For the law was given by Moses, but grace and truth came by Jesus Christ.

18 No man hath seen God at any time, the only begotten Son, which is in the bosom of the Father, he hath declared him.

19 And this is the record of John, when the Jews sent priests and Levites from Jerusalem to ask him, Who art thou?

20 And he confessed, and denied not; but confessed, I am not the Christ.

21 And they asked him, What then? Art thou Elias? And he saith, I am not. Art thou that prophet? And he answered, No.

22 Then said they unto him, Who art thou? that we may give an answer to them that sent us. What sayest thou of thyself?

23 He said, I am the voice of one crying in the wilderness, Make straight the way of the Lord, as said the prophet Esaias.

24 And they which were sent were of the Pharisees.

25 And they asked him, and said unto him, Why baptizest thou then, if thou be not that Christ, nor Elias, neither that prophet?

26 John answered them, saying, I baptize with water: but there standeth one among you, whom ye know not;

27 He it is, who coming after me is preferred before me, whose shoe's latchet I am not worthy to unloose.

28 These things were done in Bethabara beyond Jordan, where John was baptizing.

29 The next day John seeth Jesus coming unto him, and saith, Behold the Lamb of God, which taketh away the sin of the world.

30 This is he of whom I said, After me cometh a man which is preferred before me: for he was before me.

31 And I knew him not: but that he should be made manifest to Israel, therefore am I come baptizing with water.

32 And John bare record, saying, I saw the Spirit descending from heaven like a dove, and it abode upon him.

33 And I knew him not: but he that sent me to baptize with water, the same said unto me, Upon whom thou shalt see the Spirit descending, and remaining on him, the same is he which baptizeth with the Holy Ghost.

34 And I saw, and bare record that this is the Son of God.

35 Again the next day after John stood, and two of his disciples;

36 And looking upon Jesus as he walked, he saith, Behold the Lamb of God!

37 And the two disciples heard him speak, and they followed Jesus.

38 Then Jesus turned, and saw them following, and saith unto them, What seek ye? They said unto him, Rabbi, (which is to say, being interpreted, Master,) where dwellest thou?

39 He saith unto them, Come and see. They came and saw where he dwelt, and abode with him that day: for it was about the tenth hour.

40 One of the two which heard John speak, and followed him, was Andrew, Simon Peter's brother.

41 He first findeth his own brother Simon, and saith unto him, We have found the Messias, which is, being interpreted, the Christ.

42 And he brought him to Jesus. And when Jesus beheld him, he said, Thou art Simon the son of Jona: thou shalt be called Cephas, which is by interpretation, A stone.

43 The day following Jesus would go forth into Galilee, and findeth Philip, and saith unto him, Follow me.

44 Now Philip was of Bethsaida, the city of Andrew and Peter.

45 Philip findeth Nathanael, and saith unto him, We have found him, of whom Moses in the law, and the prophets, did write, Jesus of Nazareth, the son of Joseph.

46 And Nathanael said unto him, Can there any good thing come out of Nazareth? Philip saith unto him, Come and see.

47 Jesus saw Nathanael coming to him, and saith of him, Behold an Israelite indeed, in whom is no guile!

48 Nathanael saith unto him, Whence knowest thou me? Jesus answered and said unto him, Before that Philip called thee, when thou wast under the fig tree, I saw thee.

49 Nathanael answered and saith unto him, Rabbi, thou art the Son of God; thou art the King of Israel.

50 Jesus answered and said unto him, Because I said unto thee, I saw thee under the fig tree, believest thou? thou shalt see greater things than these.

51 And he saith unto him, Verily, verily, I say unto you, Hereafter ye shall see heaven open, and the angels of God ascending and descending upon the Son of man.
"@

$PlainText = @"
1 Forasmuch as many have taken in hand to set forth in order a declaration of those things which are most surely believed among us,

2 Even as they delivered them unto us, which from the beginning were eyewitnesses, and ministers of the word;

3 It seemed good to me also, having had perfect understanding of all things from the very first, to write unto thee in order, most excellent Theophilus,

4 That thou mightest know the certainty of those things, wherein thou hast been instructed.

5 There was in the days of Herod, the king of Judaea, a certain priest named Zacharias, of the course of Abia: and his wife was of the daughters of Aaron, and her name was Elisabeth.

6 And they were both righteous before God, walking in all the commandments and ordinances of the Lord blameless.

7 And they had no child, because that Elisabeth was barren, and they both were now well stricken in years.

8 And it came to pass, that while he executed the priest's office before God in the order of his course,

9 According to the custom of the priest's office, his lot was to burn incense when he went into the temple of the Lord.

10 And the whole multitude of the people were praying without at the time of incense.

11 And there appeared unto him an angel of the Lord standing on the right side of the altar of incense.

12 And when Zacharias saw him, he was troubled, and fear fell upon him.

13 But the angel said unto him, Fear not, Zacharias: for thy prayer is heard; and thy wife Elisabeth shall bear thee a son, and thou shalt call his name John.

14 And thou shalt have joy and gladness; and many shall rejoice at his birth.

15 For he shall be great in the sight of the Lord, and shall drink neither wine nor strong drink; and he shall be filled with the Holy Ghost, even from his mother's womb.

16 And many of the children of Israel shall he turn to the Lord their God.

17 And he shall go before him in the spirit and power of Elias, to turn the hearts of the fathers to the children, and the disobedient to the wisdom of the just; to make ready a people prepared for the Lord.

18 And Zacharias said unto the angel, Whereby shall I know this? for I am an old man, and my wife well stricken in years.

19 And the angel answering said unto him, I am Gabriel, that stand in the presence of God; and am sent to speak unto thee, and to shew thee these glad tidings.

20 And, behold, thou shalt be dumb, and not able to speak, until the day that these things shall be performed, because thou believest not my words, which shall be fulfilled in their season.

21 And the people waited for Zacharias, and marvelled that he tarried so long in the temple.

22 And when he came out, he could not speak unto them: and they perceived that he had seen a vision in the temple: for he beckoned unto them, and remained speechless.

23 And it came to pass, that, as soon as the days of his ministration were accomplished, he departed to his own house.

24 And after those days his wife Elisabeth conceived, and hid herself five months, saying,

25 Thus hath the Lord dealt with me in the days wherein he looked on me, to take away my reproach among men.

26 And in the sixth month the angel Gabriel was sent from God unto a city of Galilee, named Nazareth,

27 To a virgin espoused to a man whose name was Joseph, of the house of David; and the virgin's name was Mary.

28 And the angel came in unto her, and said, Hail, thou that art highly favoured, the Lord is with thee: blessed art thou among women.

29 And when she saw him, she was troubled at his saying, and cast in her mind what manner of salutation this should be.

30 And the angel said unto her, Fear not, Mary: for thou hast found favour with God.

31 And, behold, thou shalt conceive in thy womb, and bring forth a son, and shalt call his name Jesus.

32 He shall be great, and shall be called the Son of the Highest: and the Lord God shall give unto him the throne of his father David:

33 And he shall reign over the house of Jacob for ever; and of his kingdom there shall be no end.

34 Then said Mary unto the angel, How shall this be, seeing I know not a man?

35 And the angel answered and said unto her, The Holy Ghost shall come upon thee, and the power of the Highest shall overshadow thee: therefore also that holy thing which shall be born of thee shall be called the Son of God.

36 And, behold, thy cousin Elisabeth, she hath also conceived a son in her old age: and this is the sixth month with her, who was called barren.

37 For with God nothing shall be impossible.

38 And Mary said, Behold the handmaid of the Lord; be it unto me according to thy word. And the angel departed from her.

39 And Mary arose in those days, and went into the hill country with haste, into a city of Juda;

40 And entered into the house of Zacharias, and saluted Elisabeth.

41 And it came to pass, that, when Elisabeth heard the salutation of Mary, the babe leaped in her womb; and Elisabeth was filled with the Holy Ghost:

42 And she spake out with a loud voice, and said, Blessed art thou among women, and blessed is the fruit of thy womb.

43 And whence is this to me, that the mother of my Lord should come to me?

44 For, lo, as soon as the voice of thy salutation sounded in mine ears, the babe leaped in my womb for joy.

45 And blessed is she that believed: for there shall be a performance of those things which were told her from the Lord.

46 And Mary said, My soul doth magnify the Lord,

47 And my spirit hath rejoiced in God my Saviour.

48 For he hath regarded the low estate of his handmaiden: for, behold, from henceforth all generations shall call me blessed.

49 For he that is mighty hath done to me great things; and holy is his name.

50 And his mercy is on them that fear him from generation to generation.

51 He hath shewed strength with his arm; he hath scattered the proud in the imagination of their hearts.

52 He hath put down the mighty from their seats, and exalted them of low degree.

53 He hath filled the hungry with good things; and the rich he hath sent empty away.

54 He hath helped his servant Israel, in remembrance of his mercy;

55 As he spake to our fathers, to Abraham, and to his seed for ever.

56 And Mary abode with her about three months, and returned to her own house.

57 Now Elisabeth's full time came that she should be delivered; and she brought forth a son.

58 And her neighbours and her cousins heard how the Lord had shewed great mercy upon her; and they rejoiced with her.

59 And it came to pass, that on the eighth day they came to circumcise the child; and they called him Zacharias, after the name of his father.

60 And his mother answered and said, Not so; but he shall be called John.

61 And they said unto her, There is none of thy kindred that is called by this name.

62 And they made signs to his father, how he would have him called.

63 And he asked for a writing table, and wrote, saying, His name is John. And they marvelled all.

64 And his mouth was opened immediately, and his tongue loosed, and he spake, and praised God.

65 And fear came on all that dwelt round about them: and all these sayings were noised abroad throughout all the hill country of Judaea.

66 And all they that heard them laid them up in their hearts, saying, What manner of child shall this be! And the hand of the Lord was with him.

67 And his father Zacharias was filled with the Holy Ghost, and prophesied, saying,

68 Blessed be the Lord God of Israel; for he hath visited and redeemed his people,

69 And hath raised up an horn of salvation for us in the house of his servant David;

70 As he spake by the mouth of his holy prophets, which have been since the world began:

71 That we should be saved from our enemies, and from the hand of all that hate us;

72 To perform the mercy promised to our fathers, and to remember his holy covenant;

73 The oath which he sware to our father Abraham,

74 That he would grant unto us, that we being delivered out of the hand of our enemies might serve him without fear,

75 In holiness and righteousness before him, all the days of our life.

76 And thou, child, shalt be called the prophet of the Highest: for thou shalt go before the face of the Lord to prepare his ways;

77 To give knowledge of salvation unto his people by the remission of their sins,

78 Through the tender mercy of our God; whereby the dayspring from on high hath visited us,

79 To give light to them that sit in darkness and in the shadow of death, to guide our feet into the way of peace.

80 And the child grew, and waxed strong in spirit, and was in the deserts till the day of his shewing unto Israel.
"@

$allText = ""
$PlainText -split "[`r`n]+" | where {-not [string]::IsNullOrWhiteSpace($_)} | foreach {
   $line = $_
   if ($line -match "^\d+\s+(.*)$") {
       $subLine = $Matches[1] -replace "\s+",'' -replace "[^a-zA-Z]",''
       $allText += ($subLine.ToUpper() -replace "J",'I')
   }
}

$allText

Write-Host

$Matrix = @"
H K O S M
U B A V E
T W F N I
P R Z Q C
D G X L Y
"@

$alphabet = "ABCDEFGHIKLMNOPQRSTUVWXYZ"

$store = @{}

for ($i = 0; $i -lt $alphabet.Length; $i++) {
    $letter = $alphabet[$i]
    $weight = Get-Random -Minimum 0 -Maximum 10000000
    $store[$letter] = $weight
}

$newSquare = $store.GetEnumerator() | foreach {
    New-Object PSObject |
        Add-Member NoteProperty Letter $_.Key   -PassThru |
        Add-Member NoteProperty Weight $_.Value -PassThru
} | Sort-Object -Property Weight | foreach {$_.Letter}

$smatrix = $Matrix -split "[`r`n]+" -replace "\s+",''

$smatrix = @($newSquare[0..4],$newSquare[5..9],$newSquare[10..14],$newSquare[15..19],$newSquare[20..24])

$lookup = @{}
for ($row = 0; $row -lt 5; $row++) {
    $line = $smatrix[$row]
    for ($col = 0; $col -lt 5; $col++) {
        $ch = $line[$col]
        Write-Host -NoNewline "${ch} "
        $lookup[$ch] = New-Object PSObject @{Row = $row; Col = $col}
    }
    Write-Host ""
}

Write-Host

function Encode-Text ([string]$ClearText) {
    $CipherText = ""
    if ($ClearText.Length % 2 -eq 1) {
        $ClearText += 'Z'
    }
    $MsgLen = $ClearText.Length
    for ($ix = 0; $ix -lt $MsgLen; $ix += 2) {
       $ch1 = $ClearText[$ix]
       $ch2 = $ClearText[$ix+1]
       $rc1 = $lookup[$ch1]
       $rc2 = $lookup[$ch2]
       #
       # for ciphertext1 look up the character above cleartext 2
       $row = $rc2.Row - 1
       $col = $rc2.Col
       if ($row -lt 0) {$row += 5}
       $ch = ($smatrix[$row])[$col]
       $CipherText += $ch
       #
       # for ciphertext1 look up the character above cleartext 2
       $row = $rc1.Row - 1
       $col = $rc1.Col
       if ($row -lt 0) {$row += 5}
       $ch = ($smatrix[$row])[$col]
       $CipherText += $ch
    }
    $CipherText
}

Encode-Text "ARTHUR"

$cipherText = Encode-Text $allText

function Print-Chars ($AnyText, $MaxText=200) {
    for ($ix = 0; $ix -lt $MaxText; $ix += 2) {
        Write-Host -NoNewline "$($AnyText[$ix])$($AnyText[$ix+1]) "
        if ($ix % 20 -eq 18) {
            Write-Host
        }
    }

    Write-Host

    for ($ix = 0; $ix -lt $MaxText; $ix += 5) {
        Write-Host -NoNewline "$($AnyText[$ix])$($AnyText[$ix+1])$($AnyText[$ix+2])$($AnyText[$ix+3])$($AnyText[$ix+4]) "
        if ($ix % 20 -eq 15) {
            Write-Host
        }
    }
}

Write-Host

Print-Chars $allText -MaxText 400

Write-Host

Print-Chars $cipherText -MaxText 400
