# Wintropy

A fast and efficient way to save and restore windows position

## Features

- Save and restore windows position
- Launch apps as administrator
- Support for multiple monitors and virtual desktops

## Installation

Run it directly or install it as a Windows Task running at log-on

## Command Line Reference

  ```text
  Usage:        wintropy [/log=<lev>] [/admin]

  Options:

  /log=<lev>            0) off
                        1) only errors
                        2) errors and warnings
                        3) errors, warnings and info (default if omitted)
                        4) errors, warnings, info and debug
                        5) errors, warnings, info, debug and trace (everything)

  /admin                forces the app to run as administrator

  Examples:

  wintropy /log=2 /admin
  ```

***

## Disclaimer

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO EVENT SHALL THE COPYRIGHT HOLDERS OR ANYONE DISTRIBUTING THE SOFTWARE BE LIABLE FOR ANY DAMAGES OR OTHER LIABILITY, WHETHER IN CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
