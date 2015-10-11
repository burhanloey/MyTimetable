# MyTimetable
Simplified version of undergraduate class timetable for Faculty of Computer Science and Information Technology, University of Malaya.

### Instructions
MyTimetable simplifies based on the official timetable, which can be obtained from the official website.

There are two ways to load the original timetable file:

1. **As soon as the app is loaded**. You can specify which file to be loaded in ```scripts/main.js```.
2. **Drag and drop the file**. This feature is disabled by default. To enable it, remove ```display: none;``` from ```#drop``` in ```stylesheets/style.css```.

### Take caution with file name
Depends on the operating system of the server, a file name that consists of blank space may not behave as expected. If the file is not being loaded, try removing the blank spaces from the file name.

### Deployment
It is suggested to compress the JavaScript and CSS files before they are deployed. You can use sites like [JSCompress.com](http://jscompress.com/), [CSSCompressor.com](http://csscompressor.com/), or any uglifier implementation.

### Libraries
- jQuery
- Bootstrap
- KnockoutJS
- SheetJS
- MomentJS

### License
```
The MIT License (MIT)

Copyright (c) 2015 Burhanuddin Baharuddin

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```
