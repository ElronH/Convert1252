function Convertto-Win1252
{
    [CmdletBinding()]
    param
    (
        [Parameter(ValueFromPipelineByPropertyName)]
        [PSObject]$fullname
    )


    <#
    Converting from Unicode to Win1252 will break with certain extended characters.
    If the file actually uses these characters this will show false positive on converting.
    See  https://www.i18nqa.com/debug/utf8-debug.html
    #>
    $badchars = @([char] 0x00cb, [char] 0x00c2, [char] 0x00c3, [char] 0x00c5, [char] 0x00c6, [char] 0x00e2)

    $ea = [System.Text.Encoding]::GetEncoding(1252)

    [System.Collections.ArrayList]$bytesoffile = [System.IO.File]::ReadAllBytes($fullname)
    $firstbyte = $bytesoffile[0]

    if (($firstbyte -eq 254) ) # Unicode (Big-Endian)
    {
        [System.Collections.ArrayList]$a = [System.Text.Encoding]::Convert([System.Text.Encoding]::GetEncoding(1201), $ea, $bytesoffile);
        $a.removeat(0)
        [System.IO.File]::WriteAllLines($fullname, $ea.getstring($a) , $ea)

    }
    elseif (($firstbyte -eq 255) ) # Unicode (UTF-7 Little-Endian)
    {
        [System.Collections.ArrayList]$a = [System.Text.Encoding]::Convert([System.Text.Encoding]::GetEncoding(1200), $ea, $bytesoffile);
        $a.removeat(0)
        [System.IO.File]::WriteAllLines($fullname, $ea.getstring($a) , $ea)
    }
    elseif (($firstbyte -eq 239 )) # Unicode (UTF-8-BOM)
    {
        [System.Collections.ArrayList]$a = [System.Text.Encoding]::Convert([System.Text.Encoding]::GetEncoding(65001), $ea, $bytesoffile);
        $a.removeat(0)
        [System.IO.File]::WriteAllLines($fullname, $ea.getstring($a) , $ea)
    }

    else # Non Bit Correcting
    {
        $badlinez = $false
        [System.Collections.ArrayList]$linesfile = [System.IO.File]::ReadAllLines(($fullname), $ea)
        foreach ($badchar in $badchars)
        {
            if ([string]$linesfile -cmatch $badchar)
            {

           $badlinez = $true
            }

        }
        if ($badlinez -eq $true) # UTF8 NO BOM (AKA UTF7)
        {
            [System.Collections.ArrayList]$a = [System.Text.Encoding]::Convert([System.Text.Encoding]::UTF8, $ea, $bytesoffile);
            [System.IO.File]::WriteAllLines($fullname, $ea.getstring($a) , $ea)
        }
    }

    ## Removes Microsoft office 'smart' quotes and bullet points feel free to drop if you want them.
    $textoffile = [System.IO.File]::ReadAllText($fullname, $ea)
    $textoffile = $textoffile.replace(([char] 0x2018), "'")
    $textoffile = $textoffile.replace(([char] 0x2019), "'")
    $textoffile = $textoffile.replace(([char] 0x201c), "`"")
    $textoffile = $textoffile.replace(([char] 0x201D), "`"")
    $textoffile = $textoffile.replace(([char] 0x2013), "-")
    $textoffile = $textoffile.replace(([char] 0x2014), "-")
    $textoffile = $textoffile.replace(([char] 0x201A), ",")
    $textoffile = $textoffile.replace(([char] 0x00B7), " ")
    $textoffile = $textoffile.replace(([char] 0x2022), ".")
    [System.IO.File]::WriteAllText($fullname, $textoffile, $ea)

}




