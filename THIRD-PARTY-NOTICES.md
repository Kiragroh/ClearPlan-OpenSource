# Third-party notices

The vendor-free ClearPlan Simulator binary distribution includes the following
third-party libraries. They are not covered by ClearPlan's own copyright.

- OxyPlot and OxyPlot.Wpf 2.0.0;
  copyright (c) 2014 OxyPlot contributors;
  https://github.com/oxyplot/oxyplot
- Newtonsoft.Json 13.0.1;
  copyright (c) 2007 James Newton-King;
  https://github.com/JamesNK/Newtonsoft.Json
- PDFsharp and MigraDoc 1.50.5147;
  PDFsharp copyright (c) 2005-2019 empira Software GmbH;
  MigraDoc copyright (c) 2001-2019 empira Software GmbH;
  https://github.com/empira/PDFsharp

Each of these libraries is distributed under the MIT License:

> Permission is hereby granted, free of charge, to any person obtaining a copy
> of this software and associated documentation files (the "Software"), to deal
> in the Software without restriction, including without limitation the rights
> to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
> copies of the Software, and to permit persons to whom the Software is
> furnished to do so, subject to the following conditions:
>
> The above copyright notice and this permission notice shall be included in
> all copies or substantial portions of the Software.
>
> THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
> IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
> FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
> AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
> LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
> OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
> SOFTWARE.

The Eclipse/ESAPI mode additionally depends on locally supplied licensed vendor
assemblies and EsapiEssentials. Those components are not included in the
public simulator archive.

## Optional RTPLAN adapter

Distributions containing `ClearPlan.Dicom` additionally include fo-dicom 5.2.6,
copyright (c) fo-dicom contributors 2012-2026, under the Microsoft Public License
(MS-PL). Its complete upstream license and attribution notices are retained in
`licenses/fo-dicom-5.2.6-LICENSE.txt` and must accompany such distributions.
The official project is https://github.com/fo-dicom/fo-dicom.

Its managed dependency set includes CommunityToolkit.HighPerformance,
copyright (c) .NET Foundation and contributors, and Microsoft .NET libraries,
copyright (c) .NET Foundation and contributors / Microsoft Corporation.
These packages use the MIT license reproduced above. The exact versioned
inventory is `docs/DICOM_DEPENDENCIES.md`; package content hashes are pinned in
the adapter and adapter-test `packages.lock.json` files. Native image codec
packages are not included. Existing ClearPlan and vendor licensing is unchanged.
