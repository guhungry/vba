# vba
VBA modules for Microsoft Excel

## Installation
Import files as Module in `Developer` > `Visual Basic`

## Usage

### WCRegEx.Match()
Basic Regular Expression Matcher

| NAME    | TYPE   | REQUIRED | DESCRIPTION                                                             |
|---------|--------|----------|-------------------------------------------------------------------------|
| text    | String | No       | String to match                                                         |
| pattern | String | No       | Regular Expression pattern support only \d, \w, \s, [], [^], ?, + and * |

```vba
WCRegEx.Match("* Last Update 12 February 2019.", "\d+ \w+ \d\d\d\d") ' 12 February 2019
WCRegEx.Match("* Last Update 12/02/2019.", "\d\d/\d\d/\d\d\d\d") ' 12/02/2019
```

### WCString.IsSubString()
Returns True if `search` is substring of `text` else return False

| NAME   | TYPE   | REQUIRED | DESCRIPTION |
|--------|--------|----------|-------------|
| text   | String | No       | Main string |
| search | String | No       | Sub string  |

```vba
WCString.IsSubString("LONG LONG MAN", "LONG") ' True
WCString.IsSubString("LONG LONG MAN.", "SHORT") ' False
```

### WCString.IsEndsWith()
Returns True if `text` ends with `search` else return False

| NAME   | TYPE   | REQUIRED | DESCRIPTION |
|--------|--------|----------|-------------|
| text   | String | No       | Main string |
| search | String | No       | Sub string  |

```vba
WCString.IsStartsWith("LONG SHORT MAN.", "MAN") ' True
WCString.IsStartsWith("LONG SHORT MAN", "LONG") ' False
```


### WCString.IsStartsWith()
Returns True if `text` starts with `search` else return False

| NAME   | TYPE   | REQUIRED | DESCRIPTION |
|--------|--------|----------|-------------|
| text   | String | No       | Main string |
| search | String | No       | Sub string  |

```vba
WCString.IsStartsWith("LONG SHORT MAN", "LONG") ' True
WCString.IsStartsWith("LONG SHORT MAN.", "MAN") ' False
```
