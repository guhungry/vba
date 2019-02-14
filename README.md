# vba
VBA modules for Microsoft Excel

## Installation
Import files as Module in `Developer` > `Visual Basic`

## Usage

### WCRegEx.Match()
Basic Regular Expression Matcher

| NAME    | TYPE   | REQUIRED  | DESCRIPTION                |
|---------|--------|-----------|----------------------------|
| text    | String | Yes       | String to match            |
| pattern | String | Yes       | Regular Expression pattern |

`pattern` supports only \d, \w, \s, ^, $, [], [^], ?, + and *

```vba
WCRegEx.Match("* Last Update 12 February 2019.", "\d+ \w+ \d\d\d\d") ' 12 February 2019
WCRegEx.Match("* Last Update 12/02/2019.", "\d\d/\d\d/\d\d\d\d") ' 12/02/2019
```

### WCRegEx.IsMatch()
Basic Regular Expression Tester
Returns True if `text` match `pattern`'s regular expression else return False

| NAME    | TYPE   | REQUIRED  | DESCRIPTION                |
|---------|--------|-----------|----------------------------|
| text    | String | Yes       | String to test             |
| pattern | String | Yes       | Regular Expression pattern |

```vba
WCRegEx.IsMatch("* Last Update 12 February 2019.", "\d+ \w+ \d\d\d\d") ' True
WCRegEx.IsMatch("* Last Update 12-02-2019.", "\d\d/\d\d/\d\d\d\d") ' False
```

### WCString.IsSubString()
Returns True if `search` is substring of `text` else return False

| NAME   | TYPE   | REQUIRED  | DESCRIPTION |
|--------|--------|-----------|-------------|
| text   | String | Yes       | Main string |
| search | String | Yes       | Sub string  |

```vba
WCString.IsSubString("LONG LONG MAN", "LONG") ' True
WCString.IsSubString("LONG LONG MAN.", "SHORT") ' False
```

### WCString.IsEndsWith()
Returns True if `text` ends with `search` else return False

| NAME   | TYPE   | REQUIRED  | DESCRIPTION |
|--------|--------|-----------|-------------|
| text   | String | Yes       | Main string |
| search | String | Yes       | Sub string  |

```vba
WCString.IsEndsWith("LONG SHORT MAN.", "MAN") ' True
WCString.IsEndsWith("LONG SHORT MAN", "LONG") ' False
```


### WCString.IsStartsWith()
Returns True if `text` starts with `search` else return False

| NAME   | TYPE   | REQUIRED  | DESCRIPTION |
|--------|--------|-----------|-------------|
| text   | String | Yes       | Main string |
| search | String | Yes       | Sub string  |

```vba
WCString.IsStartsWith("LONG SHORT MAN", "LONG") ' True
WCString.IsStartsWith("LONG SHORT MAN.", "MAN") ' False
```
