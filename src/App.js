import React, { Component } from 'react';
import './App.css';
import { ImagePicker } from 'react-file-picker'
import fileDownload from 'js-file-download';
import GithubIcon from './icons/GitHub-Mark-64px.png'
class App extends Component {
  constructor(props) {
    super(props);
    this.state = { converting: false };
  }
  componentDidMount() {
    document.addEventListener("keydown", this.handleKeyDown, false);
  }
  rgbToHex = (rgb) => {
    var hex = Number(rgb).toString(16);
    if (hex.length < 2) {
      hex = "0" + hex;
    }
    return hex;
  };
  fullColorHex = (r, g, b, a) => {
    var red = this.rgbToHex(r);
    var green = this.rgbToHex(g);
    var blue = this.rgbToHex(b);
    var alpha = this.rgbToHex(a);
    return alpha + red + green + blue;
  };
  getColorAtXY = (x, y, imageData) => {
    var index = (y * imageData.width + x) * 4;
    var red = imageData.data[index];
    var green = imageData.data[index + 1];
    var blue = imageData.data[index + 2];
    var alpha = imageData.data[index + 3];
    return this.fullColorHex(red, green, blue, alpha);
  }
  addCellToSheet = (worksheet, address, value) => {
    /* add to worksheet, overwriting a cell if it exists */
    worksheet[address] = value;

    /* find the cell range */
    var range = window.XLSX.utils.decode_range(worksheet['!ref']);
    var addr = window.XLSX.utils.decode_cell(address);

    /* extend the range to include the new cell */
    if (range.s.c > addr.c) range.s.c = addr.c;
    if (range.s.r > addr.r) range.s.r = addr.r;
    if (range.e.c < addr.c) range.e.c = addr.c;
    if (range.e.r < addr.r) range.e.r = addr.r;

    /* update range */
    worksheet['!ref'] = window.XLSX.utils.encode_range(range);
  }
  handleUpload = (base64) => {
    this.convertToExcel(base64);
  }
  handleKeyDown = (event) => {
    if (event.keyCode === 13) {
      this.downloadImage(this.state.URL)
    }
  }
  onInput = (evt) => {
    this.setState({
      URL: evt.target.value
    })
  }
  toDataURL = (src, callback, outputFormat) => {
    this.setState({
      status: 'Downloading...'
    })
    //From https://stackoverflow.com/a/20285053
    var img = new Image();
    img.crossOrigin = 'Anonymous';
    img.onload = function () {
      var canvas = document.createElement('CANVAS');
      var ctx = canvas.getContext('2d');
      var dataURL;
      canvas.height = this.naturalHeight;
      canvas.width = this.naturalWidth;
      ctx.drawImage(this, 0, 0);
      dataURL = canvas.toDataURL(outputFormat);
      callback(dataURL);
    };
    img.src = `https://cors-anywhere.herokuapp.com/${src}`;
    if (img.complete || img.complete === undefined) {
      img.src = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==";
      img.src = src;
    }
  }
  downloadImage = url => this.toDataURL(
    url,
    base64 => {
      this.convertToExcel(base64)
    },
    'image/png'
  )


  convertToExcel = (base64) => {
    var image = new Image();
    image.src = base64;
    image.onload = () => {
      this.setState({
        status: 'Converting...'
      })
      var canvas = document.createElement('canvas');

      let destHeight = image.height;
      let destWidth = image.width;
      if (image.height > 150) {
        destHeight = 150
        destWidth = Math.floor((destHeight / image.height) * image.width)
      }
      canvas.width = destWidth;
      canvas.height = destHeight;
      var context = canvas.getContext('2d');
      context.drawImage(image, 0, 0, destWidth, destHeight);

      var imageData = context.getImageData(0, 0, canvas.width, canvas.height); //one-dimensional array of RGBA values
      var XLSX = window.XLSX;
      var url = "blank.xlsx";
      var oReq = new XMLHttpRequest();

      oReq.open("GET", url, true);
      oReq.responseType = "arraybuffer";

      oReq.onload = (e) => {
        var arraybuffer = oReq.response;

        /* convert data to binary string */
        var data = new Uint8Array(arraybuffer);
        var arr = [];
        for (var i = 0; i !== data.length; ++i) arr[i] = String.fromCharCode(data[i]);
        var bstr = arr.join("");
        var workbook = XLSX.read(bstr, { type: "binary" });
        var sheet = workbook.Sheets[workbook.SheetNames[0]]; // get the first worksheet
        /* loop through every cell */
        for (var R = 0; R < canvas.height; ++R) {
          for (var C = 0; C < canvas.width; ++C) {

            var cellref = XLSX.utils.encode_cell({ c: C, r: R }); // construct A1 reference for cell
            this.addCellToSheet(sheet, cellref, {
              t: 's',
              v: '',
              s: {
                fill: {
                  bgColor: { rgb: this.getColorAtXY(C, R, imageData) },
                  fgColor: { rgb: this.getColorAtXY(C, R, imageData) }
                }
              }

            });

          }
        }
        //Set size of rows and columns in pixels
        var wscols = Array(canvas.width).fill({ wpx: 16 })
        var wsrows = Array(canvas.height).fill({ hpx: 16 })
        sheet['!rows'] = wsrows;
        sheet['!cols'] = wscols;
        var wopts = { bookType: 'xlsx', bookSST: false, type: 'binary' };

        var wbout = XLSX.write(workbook, wopts);

        var s2ab = (s) => {
          var buf = new ArrayBuffer(s.length);
          var view = new Uint8Array(buf);
          for (var i = 0; i !== s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
          return buf;
        }

        //Download to local machine
        fileDownload(new Blob([s2ab(wbout)], { type: "" }), 'out.xlsx')
        this.setState({
          status: null
        })
      }

      oReq.send();



    };
  }
  handleUploadError = (error) => {
    alert(error);
  }
  render() {
    return (
      <div className="App">
        <header className="Title">Image to Excel file converter</header>
        <p>Enter URL to image</p>
        <input onInput={this.onInput} type="url" name="url" className="imageUrl"
          placeholder="https://example.com/Cool_Image.jpg"
        ></input>
        <p>Or</p>
        <ImagePicker
          extensions={['jpg', 'jpeg', 'png', 'webp', 'bmp']}
          onChange={this.handleUpload}
          dims={{ minWidth: 1, maxWidth: 3840, minHeight: 1, maxHeight: 2160 }}
          onError={this.handleUploadError}
        >
          <div className="Upload">Choose image</div>
        </ImagePicker>
        <p className="Note"><b>Note:</b> Images with higher resolution will be resized. Also you may want to adjust the cell width and height in the output document.</p>
        {this.state.status &&
          <p>{this.state.status}</p>
        }
        <a href="https://github.com/lopogo59/image_to_excel_converter" className="Source"><img className="sourceIcon" src={GithubIcon}/>Source code</a>
      </div>
    );
  }
}

export default App;
