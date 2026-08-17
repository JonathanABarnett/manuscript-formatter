import { webBridge } from './bridge.js';
import { init } from '../renderer/renderer.js';

/**
 * Web entry. The bridge must be in place before the UI starts, since the
 * renderer reads `window.formatter.platform` as it wires itself up.
 */
window.formatter = webBridge;
init();
