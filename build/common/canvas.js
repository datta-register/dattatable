"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.CanvasForm = void 0;
var gd_sprest_bs_1 = require("gd-sprest-bs");
/**
 * Canvas Form
 */
var _CanvasForm = /** @class */ (function () {
    // Constructor
    function _CanvasForm() {
        this._canvas = null;
        this._elBody = null;
        this._elHeader = null;
        // Render the canvas
        this.render();
    }
    Object.defineProperty(_CanvasForm.prototype, "el", {
        // Element
        get: function () { return this._canvas.el; },
        enumerable: false,
        configurable: true
    });
    // Hides the canvas
    _CanvasForm.prototype.hide = function () { this._canvas.hide(); };
    // Renders the canvas
    _CanvasForm.prototype.render = function () {
        var _this = this;
        // Create the element
        var el = document.createElement("div");
        el.id = "core-canvas";
        document.body.appendChild(el);
        // Render the canvas
        this._canvas = gd_sprest_bs_1.Components.Offcanvas({
            el: el,
            options: {
                autoClose: true,
                backdrop: true,
                focus: true,
                keyboard: true,
                scroll: true
            },
            onRenderBody: function (el) { _this._elBody = el; },
            onRenderHeader: function (el) { _this._elHeader = el; }
        });
    };
    // Sets the body
    _CanvasForm.prototype.setBody = function (content) {
        // Clear the body
        while (this._elBody.firstChild) {
            this._elBody.removeChild(this._elBody.firstChild);
        }
        // See if content exists
        if (content) {
            // See if this is text
            if (typeof (content) == "string") {
                // Set the html
                this._elBody.innerHTML = content;
            }
            else {
                // Append the element
                this._elBody.appendChild(content);
            }
        }
    };
    // Sets the header
    _CanvasForm.prototype.setHeader = function (content) {
        // Clear the body
        while (this._elHeader.firstChild) {
            this._elHeader.removeChild(this._elHeader.firstChild);
        }
        // See if content exists
        if (content) {
            // See if this is text
            if (typeof (content) == "string") {
                // Set the html
                this._elHeader.innerHTML = content;
            }
            else {
                // Append the element
                this._elHeader.appendChild(content);
            }
        }
    };
    // Sets the modal type
    _CanvasForm.prototype.setType = function (type) { this._canvas.setType(type); };
    // Shows the canvas
    _CanvasForm.prototype.show = function () { this._canvas.show(); };
    return _CanvasForm;
}());
exports.CanvasForm = new _CanvasForm();
