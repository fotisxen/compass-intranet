import * as React from 'react';
import * as ReactDom from 'react-dom';
import HelloWorld from '../src/webparts/helloWorld/components/HelloWorld';
import ChromeRoot from '../src/extensions/compassChrome/components/ChromeRoot';
import Footer from '../src/shared/components/Footer';

ReactDom.render(React.createElement(ChromeRoot, { chatApiUrl: '' }), document.getElementById('chrome-top'));
ReactDom.render(React.createElement(HelloWorld), document.getElementById('root'));
ReactDom.render(React.createElement(Footer, { companyName: 'Compass' }), document.getElementById('chrome-bottom'));
