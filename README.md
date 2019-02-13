# vba
VBA modules for Microsoft Excel

## Installation
Import files as Module in `Developer` > `Visual Basic`

## Usage

### WCRegEx.Match()
Basic Regular Expression Matcher

| NAME    | TYPE   | REQUIRED | DESCRIPTION                                             |
|---------|--------|----------|---------------------------------------------------------|
| text    | String | No       | String to match                                         |
| pattern | String | No       | Regular Expression pattern support only \d, \w, + and * |

### WCString.IsSubString()
Returns True if `search` is substring of `text` else return False

| NAME   | TYPE   | REQUIRED | DESCRIPTION |
|--------|--------|----------|-------------|
| text   | String | No       | Main string |
| search | String | No       | Sub string  |
