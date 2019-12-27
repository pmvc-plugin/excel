<?php

namespace PMVC\PlugIn\auth;

${_INIT_CONFIG
}[_CLASS] = __NAMESPACE__.'\XLS_PARSE';

class XLS_PARSE
{
    function v($data,$pos) {
      return ord($data[$pos]) | ord($data[$pos+1])<<8;
    }

    function myHex($d) 
    {
        if ($d < 16) { return "0" . dechex($d);
        }
        return dechex($d);
    }
    
    function dumpHexData($data, $pos, $length) 
    {
        $info = "";
        for ($i = 0; $i <= $length; $i++) {
            $info .= ($i==0?"":" ") . $this->myHex(ord($data[$pos + $i])) . (ord($data[$pos + $i])>31? "[" . $data[$pos + $i] . "]":'');
        }
        return $info;
    }

    function _GetInt4d($data, $pos) 
    {
        $value = ord($data[$pos]) | (ord($data[$pos+1]) << 8) | (ord($data[$pos+2]) << 16) | (ord($data[$pos+3]) << 24);
        if ($value>=4294967294) {
            $value=-2;
        }
        return $value;
    }

    public function __invoke($version, $data, $code, $pos, $length, $lineStyles, $dateFormats, $numberFormats)
    {
        $sst = [];
        $formatRecords = [];
        $fontRecords = [];
        $thisColors = [];
        $xfRecords = [];
        $boundsheets = [];

        while ($code != SPREADSHEET_EXCEL_READER_TYPE_EOF) {
            switch ($code) {
            case SPREADSHEET_EXCEL_READER_TYPE_SST:
                $spos = $pos + 4;
                $limitpos = $spos + $length;
                $uniqueStrings = $this->_GetInt4d($data, $spos+4);
                $spos += 8;
                for ($i = 0; $i < $uniqueStrings; $i++) {
                    // Read in the number of characters
                    if ($spos == $limitpos) {
                        $opcode =$this->v($data, $spos);
                         
                        $conlength = $this->v($data, $spos+2);
                        if ($opcode != 0x3c) {
                          break;
                        }
                        $spos += 4;
                        $limitpos = $spos + $conlength;
                    }
                    $numChars = ord($data[$spos]) | (ord($data[$spos+1]) << 8);
                    $spos += 2;
                    $optionFlags = ord($data[$spos]);
                    $spos++;
                    $asciiEncoding = (($optionFlags & 0x01) == 0) ;
                    $extendedString = ( ($optionFlags & 0x04) != 0);

                    // See if string contains formatting information
                    $richString = ( ($optionFlags & 0x08) != 0);

                    if ($richString) {
                        // Read in the crun
                        $formattingRuns = $this->v($data, $spos);
                        $spos += 2;
                    }

                    if ($extendedString) {
                        // Read in cchExtRst
                        $extendedRunLength = $this->_GetInt4d($data, $spos);
                        $spos += 4;
                    }

                    $len = ($asciiEncoding)? $numChars : $numChars*2;
                    if ($spos + $len < $limitpos) {
                        $retstr = substr($data, $spos, $len);
                        $spos += $len;
                    }
                    else{
                        // found countinue
                        $retstr = substr($data, $spos, $limitpos - $spos);
                        $bytesRead = $limitpos - $spos;
                        $charsLeft = $numChars - (($asciiEncoding) ? $bytesRead : ($bytesRead / 2));
                        $spos = $limitpos;

                        while ($charsLeft > 0){
                            $opcode = $this->v($data, $spos);
                            $conlength = $this->v($data, $spos+2);
                            if ($opcode != 0x3c) {
                              break;
                            }
                            $spos += 4;
                            $limitpos = $spos + $conlength;
                            $option = ord($data[$spos]);
                            $spos += 1;
                            if ($asciiEncoding && ($option == 0)) {
                                $len = min($charsLeft, $limitpos - $spos); // min($charsLeft, $conlength);
                                $retstr .= substr($data, $spos, $len);
                                $charsLeft -= $len;
                                $asciiEncoding = true;
                            }
                            elseif (!$asciiEncoding && ($option != 0)) {
                                $len = min($charsLeft * 2, $limitpos - $spos); // min($charsLeft, $conlength);
                                $retstr .= substr($data, $spos, $len);
                                $charsLeft -= $len/2;
                                $asciiEncoding = false;
                            }
                            elseif (!$asciiEncoding && ($option == 0)) {
                                // Bummer - the string starts off as Unicode, but after the
                                // continuation it is in straightforward ASCII encoding
                                $len = min($charsLeft, $limitpos - $spos); // min($charsLeft, $conlength);
                                for ($j = 0; $j < $len; $j++) {
                                    $retstr .= $data[$spos + $j].chr(0);
                                }
                                $charsLeft -= $len;
                                $asciiEncoding = false;
                            }
                            else{
                                $newstr = '';
                                for ($j = 0; $j < strlen($retstr); $j++) {
                                    $newstr = $retstr[$j].chr(0);
                                }
                                $retstr = $newstr;
                                $len = min($charsLeft * 2, $limitpos - $spos); // min($charsLeft, $conlength);
                                $retstr .= substr($data, $spos, $len);
                                $charsLeft -= $len/2;
                                $asciiEncoding = false;
                            }
                            $spos += $len;
                        }
                    }
                    $retstr = ($asciiEncoding) ? $retstr : \PMVC\plug('utf8')->decodeUtf16($retstr);

                    if ($richString) {
                        $spos += 4 * $formattingRuns;
                    }

                    // For extended strings, skip over the extended string data
                    if ($extendedString) {
                        $spos += $extendedRunLength;
                    }
                    $sst[]=$retstr;
                }
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_FILEPASS:
                return false;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_NAME:
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_FORMAT:
                $indexCode = $this->v($data, $pos+4);
                if ($version == SPREADSHEET_EXCEL_READER_BIFF8) {
                    $numchars = $this->v($data, $pos+6);
                    if (ord($data[$pos+8]) == 0) {
                        $formatString = substr($data, $pos+9, $numchars);
                    } else {
                        $formatString = substr($data, $pos+9, $numchars*2);
                    }
                } else {
                    $numchars = ord($data[$pos+6]);
                    $formatString = substr($data, $pos+7, $numchars*2);
                }
                $formatRecords[$indexCode] = $formatString;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_FONT:
                $height = $this->v($data, $pos+4);
                $option = $this->v($data, $pos+6);
                $color  = $this->v($data, $pos+8);
                $weight = $this->v($data, $pos+10);
                $under  = ord($data[$pos+14]);
                $font = "";
                // Font name
                $numchars = ord($data[$pos+18]);
                if ((ord($data[$pos+19]) & 1) == 0) {
                    $font = substr($data, $pos+20, $numchars);
                } else {
                    $font = substr($data, $pos+20, $numchars*2);
                    $font =  \PMVC\plug('utf8')->decodeUtf16($font); 
                }
                $fontRecords[] = array(
                               'height' => $height / 20,
                               'italic' => !!($option & 2),
                               'color' => $color,
                               'under' => !($under==0),
                               'bold' => ($weight==700),
                               'font' => $font,
                               'raw' => $this->dumpHexData($data, $pos+3, $length)
                );
                break;

            case SPREADSHEET_EXCEL_READER_TYPE_PALETTE:
                $colors = ord($data[$pos+4]) | ord($data[$pos+5]) << 8;
                for ($coli = 0; $coli < $colors; $coli++) {
                    $colOff = $pos + 2 + ($coli * 4);
                    $colr = ord($data[$colOff]);
                    $colg = ord($data[$colOff+1]);
                    $colb = ord($data[$colOff+2]);
                    $thisColors[0x07 + $coli] = '#' . $this->myhex($colr) . $this->myhex($colg) . $this->myhex($colb);
                }
                break;

            case SPREADSHEET_EXCEL_READER_TYPE_XF:
                $fontIndexCode = (ord($data[$pos+4]) | ord($data[$pos+5]) << 8) - 1;
                $fontIndexCode = max(0, $fontIndexCode);
                $indexCode = ord($data[$pos+6]) | ord($data[$pos+7]) << 8;
                $alignbit = ord($data[$pos+10]) & 3;
                $bgi = (ord($data[$pos+22]) | ord($data[$pos+23]) << 8) & 0x3FFF;
                $bgcolor = ($bgi & 0x7F);
                //                        $bgcolor = ($bgi & 0x3f80) >> 7;
                $align = "";
                if ($alignbit==3) { $align="right"; 
                }
                if ($alignbit==2) { $align="center"; 
                }

                $fillPattern = (ord($data[$pos+21]) & 0xFC) >> 2;
                if ($fillPattern == 0) {
                    $bgcolor = "";
                }

                $xf = array();
                $xf['formatIndex'] = $indexCode;
                $xf['align'] = $align;
                $xf['fontIndex'] = $fontIndexCode;
                $xf['bgColor'] = $bgcolor;
                $xf['fillPattern'] = $fillPattern;

                $border = ord($data[$pos+14]) | (ord($data[$pos+15]) << 8) | (ord($data[$pos+16]) << 16) | (ord($data[$pos+17]) << 24);
                $xf['borderLeft'] = $lineStyles[($border & 0xF)];
                $xf['borderRight'] = $lineStyles[($border & 0xF0) >> 4];
                $xf['borderTop'] = $lineStyles[($border & 0xF00) >> 8];
                $xf['borderBottom'] = $lineStyles[($border & 0xF000) >> 12];
                        
                $xf['borderLeftColor'] = ($border & 0x7F0000) >> 16;
                $xf['borderRightColor'] = ($border & 0x3F800000) >> 23;
                $border = (ord($data[$pos+18]) | ord($data[$pos+19]) << 8);

                $xf['borderTopColor'] = ($border & 0x7F);
                $xf['borderBottomColor'] = ($border & 0x3F80) >> 7;
                                                
                if (array_key_exists($indexCode, $dateFormats)) {
                    $xf['type'] = 'date';
                    $xf['format'] = $dateFormats[$indexCode];
                    if ($align=='') { $xf['align'] = 'right'; 
                    }
                }elseif (array_key_exists($indexCode, $numberFormats)) {
                    $xf['type'] = 'number';
                    $xf['format'] = $numberFormats[$indexCode];
                    if ($align=='') { $xf['align'] = 'right'; 
                    }
                }else{
                    $isdate = false;
                    $formatstr = '';
                    if ($indexCode > 0) {
                        if (isset($formatRecords[$indexCode])) {
                            $formatstr = $formatRecords[$indexCode];
                        }
                        if ($formatstr!="") {
                            $tmp = preg_replace("/\;.*/", "", $formatstr);
                            $tmp = preg_replace("/^\[[^\]]*\]/", "", $tmp);
                            if (preg_match("/[^hmsday\/\-:\s\\\,AMP]/i", $tmp) == 0) { // found day and time format
                                $isdate = true;
                                $formatstr = $tmp;
                                $formatstr = str_replace(array('AM/PM','mmmm','mmm'), array('a','F','M'), $formatstr);
                                // m/mm are used for both minutes and months - oh SNAP!
                                // This mess tries to fix for that.
                                // 'm' == minutes only if following h/hh or preceding s/ss
                                $formatstr = preg_replace("/(h:?)mm?/", "$1i", $formatstr);
                                $formatstr = preg_replace("/mm?(:?s)/", "i$1", $formatstr);
                                // A single 'm' = n in PHP
                                $formatstr = preg_replace("/(^|[^m])m([^m]|$)/", '$1n$2', $formatstr);
                                $formatstr = preg_replace("/(^|[^m])m([^m]|$)/", '$1n$2', $formatstr);
                                // else it's months
                                $formatstr = str_replace('mm', 'm', $formatstr);
                                // Convert single 'd' to 'j'
                                $formatstr = preg_replace("/(^|[^d])d([^d]|$)/", '$1j$2', $formatstr);
                                $formatstr = str_replace(array('dddd','ddd','dd','yyyy','yy','hh','h'), array('l','D','d','Y','y','H','g'), $formatstr);
                                $formatstr = preg_replace("/ss?/", 's', $formatstr);
                            }
                        }
                    }
                    if ($isdate) {
                        $xf['type'] = 'date';
                        $xf['format'] = $formatstr;
                        if ($align=='') { $xf['align'] = 'right'; 
                        }
                    }else{
                        // If the format string has a 0 or # in it, we'll assume it's a number
                        if (preg_match("/[0#]/", $formatstr)) {
                            $xf['type'] = 'number';
                            if ($align=='') { $xf['align']='right'; 
                            }
                        }
                        else {
                            $xf['type'] = 'other';
                        }
                        $xf['format'] = $formatstr;
                        $xf['code'] = $indexCode;
                    }
                }
                $xfRecords[] = $xf;
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_NINETEENFOUR:
                $nineteenFour = (ord($data[$pos+4]) == 1);
                break;
            case SPREADSHEET_EXCEL_READER_TYPE_BOUNDSHEET:
                $rec_offset = $this->_GetInt4d($data, $pos+4);
                $rec_typeFlag = ord($data[$pos+8]);
                $rec_visibilityFlag = ord($data[$pos+9]);
                $rec_length = ord($data[$pos+10]);

                if ($version == SPREADSHEET_EXCEL_READER_BIFF8) {
                    $chartype =  ord($data[$pos+11]);
                    if ($chartype == 0) {
                        $rec_name    = substr($data, $pos+12, $rec_length);
                    } else {
                        $rec_name    = \PMVC\plug('utf8')->decodeUtf16(substr($data, $pos+12, $rec_length*2));
                    }
                }elseif ($version == SPREADSHEET_EXCEL_READER_BIFF7) {
                    $rec_name    = substr($data, $pos+11, $rec_length);
                }
                $boundsheets[] = array('name'=>$rec_name,'offset'=>$rec_offset);
                break;

            }
            $pos += $length + 4;
            $code = ord($data[$pos]) | ord($data[$pos+1])<<8;
            $length = ord($data[$pos+2]) | ord($data[$pos+3])<<8;
        }

        return [
          'sst'=>$sst,
          'fontRecords'=>$fontRecords,
          'colors'=>$thisColors,
          'xfRecords'=>$xfRecords,
          'boundsheets'=>$boundsheets,
          'nineteenFour'=>$nineteenFour,
        ];
    }
}
