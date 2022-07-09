<div id="top"></div>
<br />
<div align="center">

  <h3 align="center">Annual Leave System - Google Apps Script Replacement</h3>

  <p align="center">
    An awesome README template to jumpstart your projects!
    <br />
    <a href="https://github.com/jalexm8/Google-App-Scripts/annual-leave/"><strong>Explore the docs »</strong></a>
    <br />
    <br />
    <a href="https://github.com/jalexm8/Google-App-Scripts/annual-leave/">View Demo (soon!)</a>
    ·
    <a href="https://github.com/jalexm8/Google-App-Scripts/issues/">Report Bug</a>
    ·
    <a href="https://github.com/jalexm8/Google-App-Scripts/issues">Request Feature</a>
  </p>
</div>



<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li>
      <a href="#about-the-project">About The Project</a>
      <ul>
        <li><a href="#built-with">Built With</a></li>
      </ul>
    </li>
    <li>
      <a href="#getting-started">Getting Started</a>
      <ul>
        <li><a href="#prerequisites">Prerequisites</a></li>
        <li><a href="#installation">Installation</a></li>
        <ul>
          <li><a href="#`pto_form.js`">`pto_form.js`</a></li>
        </ul>
      </ul>
    </li>
    <li><a href="#usage">Usage</a></li>
    <li><a href="#roadmap">Roadmap</a></li>
  </ol>
</details>



<!-- ABOUT THE PROJECT -->
## About The Project

Can we replace an annual leave system with Google Apps Script? Lets find out!

<p align="right">(<a href="#top">back to top</a>)</p>

### Built With
* [Google Apps Script](https://developers.google.com/apps-script)

<p align="right">(<a href="#top">back to top</a>)</p>

<!-- GETTING STARTED -->
## Getting Started

### Prerequisites
* A Google account.
* You need to create:  
  * A Google Form (to act as a "PTO request form").
    1. Navigate to https://drive.google.com/drive/my-drive.
    2. Create a blank Google form.
    3. Fill in the following:
        * Title: "PTO Request Form"
        * 1st Question: "Start Date"
        * 2nd Question: "End Date"
        * Ensure email collection is enabled in settings.
  * A Google Sheet (to acts as a "database" _yuck_).
    1. Navigate to https://drive.google.com/drive/my-drive.
    2. Create a blank Google sheet.
    3. Insert the headers like: `name,	email,	line manager email,	pto allowance,	pto remaining,	pto requested,	pto authorised`

### Installation
#### `pto_form.js`
  1. Open up the "PTO request form" Google form.
  2. Press the kebab menu (top right) -> Script editor.
  3. Paste the contents of `pto_form.js` and save.

<p align="right">(<a href="#top">back to top</a>)</p>

<!-- USAGE EXAMPLES -->
## Usage
_Coming soon_

<p align="right">(<a href="#top">back to top</a>)</p>

<!-- ROADMAP -->
## Roadmap
- [ ] Create annual leave system within Google Apps Scripts

See the [open issues](https://github.com/jalexm8/Google-App-Scripts/issues) for a full list of proposed features (and known issues).

<p align="right">(<a href="#top">back to top</a>)</p>
