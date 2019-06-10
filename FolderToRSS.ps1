

# Copyright (c) 2019 Gurjit Singh

# This source code is licensed under the MIT license that can be found in
# the accompanying LICENSE file or at https://opensource.org/licenses/MIT.


# https://help.apple.com/itc/podcasts_connect/#/itcb54353390
# https://github.com/simplepie/simplepie-ng/wiki/Spec:-iTunes-Podcast-RSS


$Dir = $PSScriptRoot

$DirName = ($PSScriptRoot | Split-Path -Leaf -Resolve)

$URL = "http://10.1.1.9/$DirName/"

$OutName = $Dir + "\feed.xml"

$AudioExt = @(".mp3"; ".m4a")

$ArtworkFileName = "artwork.jpg"

$ShellDir = ((New-Object -COMObject Shell.Application).Namespace($Dir))

$EachFile = $ShellDir.items() | 
            Where-Object {
            ($PSItem.ExtendedProperty("System.FileExtension")) -in $AudioExt
            }

$CompTitle = ($EachFile | Select-Object -Index 0).ExtendedProperty("System.Music.AlbumTitle")

$Author = ($EachFile | Select-Object -Index 0).ExtendedProperty("System.Music.AlbumArtist")

if (!($ShellDir -and $EachFile -and $CompTitle -and $Author)) {

    throw
}


foreach ($file in $EachFile) {

    $TrackTitle = ($file.ExtendedProperty("System.Title"))

    $TrackAuthor = ($file.ExtendedProperty("System.Music.Artist"))

    $TrackNum = ($file.ExtendedProperty("System.Music.TrackNumber"))

    $TrackExt = ($file.ExtendedProperty("System.FileExtension"))

    $TrackLength = ($ShellDir.GetDetailsOf($file, 27))

    $filename = ($file.Path | Split-Path -Leaf -Resolve)

    $TrackType = switch ($TrackExt) {
        ".mp3" { "audio/mpeg"; break }
        ".m4a" { "audio/x-m4a"; break }
        Default { throw }
    }

    if (!($TrackTitle -and $TrackAuthor -and $TrackNum -and $TrackExt -and $TrackLength -and $filename)) {

        throw
    }

    $FakeDate = (Get-Date -Day 1 -Month 1 -Year 2009).AddDays($TrackNum)
    
    $PerTrack += @"

        <item>
            <title>$TrackTitle</title>
            <itunes:title>$TrackTitle</itunes:title>
            <enclosure url=`"$URL$filename`" type=`"$TrackType`" length=`"$($file.size)`"/>
            <guid>`"$URL$filename`"</guid>
            <itunes:author>$TrackAuthor</itunes:author>
            <itunes:duration>$TrackLength</itunes:duration>
            <itunes:explicit>false</itunes:explicit>
            <itunes:episode>$TrackNum</itunes:episode>
            <pubDate>$(Get-Date $FakeDate -UFormat "%a, %d %b %Y %T GMT")</pubDate>
        </item> 

"@

}


@"
<?xml version="1.0" encoding="UTF-8"?>
<rss version="2.0" xmlns:itunes="http://www.itunes.com/dtds/podcast-1.0.dtd" xmlns:content="http://purl.org/rss/1.0/modules/content/">
    <channel>
        <title>$CompTitle</title>
        <language>en-us</language>
        <itunes:author>$Author</itunes:author>
        <itunes:explicit>false</itunes:explicit>
        <itunes:type>serial</itunes:type>
        <itunes:image href="$URL$ArtworkFileName" />
        $PerTrack
    </channel>
</rss>
                
"@ | Out-File -FilePath $OutName
