import React, { useState, useEffect, useCallback, useRef } from "react";
import {
  ClipboardList, BarChart3, Settings, Check, Radio, Plus, Trash2, RotateCcw, List, BarChart2, Layers,
  LogOut, Lock, User, ShieldCheck, Download, Play, Pause, Timer as TimerIcon, Radar as RadarIcon, Star, PartyPopper, FileText, Upload,
} from "lucide-react";
import {
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, Cell, LabelList,
  RadarChart, PolarGrid, PolarAngleAxis, PolarRadiusAxis, Radar, Legend,
} from "recharts";
import * as XLSX from "xlsx";
import storage from "./storageClient.js";

const LOGO2_SRC = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAJQAAABRCAYAAAA0GY/TAAAAAXNSR0IArs4c6QAAAAlwSFlzAAAOxAAADsQBlSsOGwAAABl0RVh0U29mdHdhcmUATWljcm9zb2Z0IE9mZmljZX/tNXEAABAVSURBVHhe7Z0JtFXTH8d/adSglAZFKalEojK0RIpEhqyoVoZFyJhCRRRRaFgoQ0VmoTKUKGMZllkUGkSSoUmheR78z+f3/O7/vNurd++7+9x3r7V/a9317jt373323fv7fuN3n1fsn0DEi18BByswcODAIsUcjOOH8CsQWwEPKA8GpyvgAeV0Of1gHlAeA05XwAPK6XL6wTygPAacroAHlNPl9IN5QHkMOF0BDyiny+kH84DyGHC6Ah5QTpfTD+YB5THgdAU8oJwupx/MA8pjwOkKeEA5XU4/mAeUx4DTFfCAcrqcfjAPKI8BpyvgAeV0Of1g7gA1t4nIjnWZv6KlG4nUnVjgee7cuVOeeOIJ2bx5szRq1EhOOumkAo+VSMdzzjlH3nvvPbn66qtl6NChiXQp1DbuALVxVqF+kYRv/s+WhJvm1XDSpElyxRVXxD7aunWrFC9ePKUx99T5999/l3Xr1smKFSsiu4fLgd0Bqmj5QEOtcTm3aMYqXjWlcR9//PFc/d966y0566yzBM2111577XbsHTt2SNGiRfVz2m7fvl1KlCix2/bbtm1ToB588MEyc+ZMKVOmzC5tAfOexsjv85QWYjed3QEqitll2JiYOQCEmWvYsKGMGjVKxo8fr4D64IMP5JprrpF99tlH21SsWFFn37lzZ/nyyy9l4cKFsmTJErnqqqtkypQp+lmzZs0kOHokp59+euybDhs2TEaPHi2//PKLHHnkkQKwEAPryy+/LHfeeafMmTNHr1euXFkwi/fee6/eGxD16NFDnnnmGTXLBx54oPBHcOqpp6ZlNT2gklhmA0K/fv3kkEMOUUC9+uqrOkLjxo3lhx9+0PcTJ06Uyy+/XP7880958cUX5bDDDlNAHHHEEfL333+rP1S6dGm57777pF27dvL555/LscceKzfeeKMMHz5cx8A/w9z99ddf+nuxYjlb1adPHwXbueeeK7Vq1ZLnn39eHnvsMdmwYYO+7927tzz66KMKxqZNm6q/N2HCBA+oJPY5bU3HjBkjRYoUkVNOOUXvWb58eVmzZo1MmzZNr+Fb0eall15SQAEm5M0335TXXntNwXT++ecrEBE0yoABA3TT0XgGpoceeki6d++ubdBiX3/9tZpIBAe9du3a+n79+vUyf/58eeONN+S7777Ta2+//bb+RIsyHtpp+fLlaVsjr6ESXGpMybvvvqut0Q5sMFoBGTt2rALq+uuvV0BNnz5drz/99NP6E7Nzyy236Hs0GqAMyzfffCPvvPOOXsJvMjDx+/7776/XzfQBJjTSHXfcIUuXLo0Ns/fee+v7M844Q3788UcZMWKETJ06VcFrfwAJftWUmqUJUEEUVLJysFqB0/pP7sVMafa7dA5O1W8L/pK3/CFSxO0Je/wi23BMWlgmT56svx566KGqaebNmyc333yzzJgxQ7p06aKfLVu2TH/iZGMet2zZohrqgAMOUPP2xx/BnAOJd+xx5pFy5crpzzZt2qhGpC/+1OzZs9Wnsnb333+/+lEjR46UBQsWaHsiU/ysdEh6AFWmfrDaQVohHSfftwUbM7tOEHFudLp+mCUE8LBJCFoFjYDZQ3txHT8Ic4dzjdx00036s27dujFzhUaLl88++0wvATTAe9ppp+nvpAwQNBWPoQBMCE7+fvvtJ4sXL9bfASJaE1/r4YcfjvlnmMhHHnnkPwaoEqVzwDRvpMiamcHXj0BL8cyPWp1EarTN0YQ5f9hOZO3ateoDIeGIDIcaU7Np0yaNzADUhRdeqCaLCKtOnTrqHCPXXXedmkPGoV3btm1l9erVuvmA7tZbb1Xnfe7cuXqPK6+8UsHy8ccfa/9Vq1bl+i449vXq1ZPBgwfrdVISpCPQkieeeKK0b99e54Xg8KdL0qOhzPosCJzRTfOi+W7co2jJHEA5FlIC+D1sYrxceuml6vhi3jA1JUuWVF9qyJAhqqlMDj/8cHXWidLQMqZpMGW24WiTiy66SDUfkRpSv359jeoAB3MgW44jj7lD0HxEg4xDTgqQAVxepUqVko4dOypY0yXpAZT8i6iyNaMDFCtWKvDTIpATTjhBfRxyPvGChmGDMVWWZCRdgFxyySW5mp933nnCi8gMDbbvvvtq6G9SpUoVjdIWLVqkERy+FbJy5UoFB4I2u/jii3U++GMkPPncfC9MMs46zjtaM34OESxPriHTBKiov0a047PxexIDGpHYs88+qxoLzWIRWnzfBg0a7HE8SwtYo3ggV61aVXjt7vPq1avLgw8+KF27dtX8E7+nSzygHK409TYzc08++aTDkZMbCnOI+TzooINi5Z7kRih4aw+ogq/dLj3RJJgs8k44x4UlaDhKPmTI0y0eUA5XHB+qIDUzfCBSDzjWYVNW0KlRcuFVGOIBlcCqf/vtt1reQChjEMmhhVq0aCFHHXVUAiPsuQn1N3wvMvAWvaU8aB4DWPnG6oJR3MMDKoFVpRa3u9Cbgiz1uVSEiA8hUoxKyKiT+yL18Ouvv8aKza7v5wGVwIoaFYWmhOtkrMlkU4MjN4W/YlwnQnZMGOaPHFE8+Q72AAnLChUqxFIGlhKwehz3IUNO4pO0gN0fpx8KDL+H0w3kv6grWjRKchQhUYrwGdes/PP9999ryads2bLOyYEeUAkAKtyEXBHSunVrBRRZdITkY/PmzXNV9kkboNms2EuBmJIMGW2EjPaHH36o2XbECsAwB44//njVWBR6AdAFF1wgL7zwQmwqAPq5556TSpUqaRK0b9++UrNmTfnpp59izARqhl988YXSY6ymyADQaBBSC2TwXYoHVJKrSWGYjWYzqaUBGLTTRx99pJqJovBxxx2nBLjbbrtNN4wyChtN9hwe1Lhx47Q9oT1iPg2lGoTSC4lNCryE/iRDX3nlFS3rXHbZZepn8RklHNicAHHjxo2aMG3ZsqUmNWE64PuRtSfxSs0R9gECuNB8UTjuHlBJAgrH2QSnvGfPnvrr2WefHdMMmMVPP/001o6yCeURBLMEOY/2JoDHgHXXXXdpppt8FgxQzBRgQmAVADoy6ABq1qxZCiTKLQjgo0xkY5Fgff/99zVjToGYQAIJa7okv36+zT2g8l2i3A2ggkCsQ9tQc0MbQFfhGjwltJJRUYgG0WZoLrQW5RF8ICguaBeAQqrAqCfGUKAfYyGAxoRSS7z8/PPPSmVB8IlMDGT4aggm2QTARpU994BKElDGK4JFgG/z22+/qdlBy8DYxHxBosM/qlatmo7OpgMGois0D8lP6C6tWrWSr776KsZ1wrdhs6kFYjopBBtYGActQ5SGD8bYaDv8JvwkxHww3qO5ECPzhf8DC4CNSjygCriyn3zySawnGwyTAAFUUEfQTERfCP4K0R2RFVwn/DBMJ9RexEBDIRd2JcBCW1EExqyaAOIwFcX4T/HUFtrH/wufcErCithEjRZkFHAZdunmAZXASlokR1OiOwBEdIbUqFFDN934UlBuicyIrEwoBqNdONxAwdY0CXxxBGYlQr7rhhtukDPPPFNPxhD2kzqA80SEiNkk+sO8QjMGJADHtJD5aYxlIDOaMGAOzwcw3X777WqiXYoHVAKrST6pSZPgZHQgAAVAsdkcBIDbTZQHrYT8knGdCOvxjdBM/OzQoYOyPZ966ikdh+NUHKFC4EpBWSHMRziaRX/MKdEkKQHGIHIzzYgDDvgQzB4Ze4sSbUyu2ZiA+p577tExADxtifxciwdUAiuK35QfJ5saHMXYvA58AgYKx/g6bCqOPZGXCdEbLxNMJGkFxFiXaDdeZubC04ZExyss3bp1E15hQcvxCh86TeDrJ9XEAyqp5cq/cV6nhy2LTm8YmwADc4dDnp+Es+e0dVGHC88nv/sn+7kHVLIrlkJ7tA1JTswQhxn+i+IBlcZdJVzH5BQmVyrqr+sBFfUKh8bHHLoGE9Ee2XSy5+aAp/Er7XKrNAHq32NT63PKD5HJ5pWRDM2mEfYj+B+WxyFst0Jv//79C2VDMaGc0SNhatn1SBYhwUHTAyg7xVunq8hqzuVFJNX//xQTl3egLpcf8Y0CbmFoCKJLAGUlFpffuyBjpQdQW/8ljjXqVZA5Jtdnx+qgBpFDD3ElFGntgCfFVyI1BJYlJozEIpV7OEyAj6QiGWhKK5Z0hJNEzY+sOCdiILwhdlQqvi85rviIkWQq5o1kKlQZIj6L2MLPSyBRCv8J4h6AI/FqvCxyZSQ7qSdyjTm5YJ3aWqcHUBvmi8ypG5zoDR64FemzDQIg8Tyl7cGJWYeHk+Ec8ULC9TLKJHZUikw5WXCesGJC8pAnsAA2nhFFlh2x08a8J9NNOQZqij26h+v0JcF59NFHa22Pck6YwUDik8y6pREMMOSj4rUpLAMK0Yx19913q/nm+9j94ss0qfwhpgdQEtS0Ni0MXqlMNYm+DsEUf1djEnAd5gCAguvEhiNoM46oP/DAA1oApkiMdgofPqAP2oviMllr036wP8leW1+4TRR54U8BpmOOOUZrhmTcMXNhCWvCa6+9VucBk4H6ILwris38tMIwYKLEE38GMIlVzrNpmgCV6jQzs79toj3viRINPCUEkJDxBhA8CCPsX5Exh5yH9OqV4wawsTYOoIKIR94Kuq5RhDGlACnMsrSDB1aItsMU1OrgQ2HSEOOt21gA/fXXX3e+sB5QKSyp+S9oGSQMGvwfQAIgYE6G/RQ0lgHKNI3Rcm06gBM/C4DA0gRE+EaYTUwXxWP8NIsyzWxhWtFOxmQw7rll3K0dgI1CPKBSWFXbHEsj2KN3bEhjKcTzj/KiksT35ZweQoGYCA6SHiaRUzY8mAOfC4KfaRxMKvMxegsJVOqG8M0pRMdL2BdMYQl26eoBlcJqGtPy5JNP1oIvpg0QEcnxHh8LIWIzwlv87fCl8Iswg1BVAA+MBoCE4JthtojKYB5AYUFbmQYyMFN8NrNHP3wmxGg2KXzNpLq6A1Q2PFKapdm+OqkFim8c1iTGBIBUh++EeUNTABIiMIQnBGP6jD3AtbCGIiHJE1wAjTnr1hcwYeKgzkAFhvVpUaQ9Wsh8JA5FAC5MLWYQYMLoNF6UUYAtqIA9GoW4A1SJGsFDvnKOFGW0MM8UBE3RqVMnzS8ZaY2QHTPECRgeuAog2FA0iXGe8JmsX5glaX0xUQCLvhDo6Dto0CCdKced+AyuFX3hR9mT8UhVACrYnMwJ7QbfHVoy+S5yUPhilu/isCd/CFElYd0BqvHiFLYpe7qSU8rrIRREdIT7vPISEqC7e3gFDvKe+gKKMBU4PD5HpMKCw26nZPKaB0eowmf0XK+8O0C5npkfLytXwAMqK7ctcyftAZW5e5OVM/OAyspty9xJe0Bl7t5k5cw8oLJy2zJ30h5Qmbs3WTkzD6is3LbMnbQHVObuTVbOzAMqK7ctcyftAZW5e5OVM/OAyspty9xJe0Bl7t5k5cw8oLJy2zJ30h5Qmbs3WTkzD6is3LbMnXSxgFEY4Sm2zP3ifmbRrMD/AL10MK5wejAlAAAAAElFTkSuQmCC";

//---------------------------------------------------------------------
// Minimal QR encoder (Byte mode only), adapted from the public-domain
// kazuhikoarase/qrcode-generator (MIT). Runs fully offline in the
// browser — no network requests, so it works under strict CSPs.
//---------------------------------------------------------------------
const QRMode = { MODE_8BIT_BYTE: 1 << 2 };
const QRErrorCorrectionLevel = { L: 1, M: 0, Q: 3, H: 2 };
const QRMaskPattern = {
  PATTERN000: 0, PATTERN001: 1, PATTERN010: 2, PATTERN011: 3,
  PATTERN100: 4, PATTERN101: 5, PATTERN110: 6, PATTERN111: 7,
};

const QRMath = (function () {
  const EXP_TABLE = new Array(256);
  const LOG_TABLE = new Array(256);
  for (let i = 0; i < 8; i++) EXP_TABLE[i] = 1 << i;
  for (let i = 8; i < 256; i++) {
    EXP_TABLE[i] = EXP_TABLE[i - 4] ^ EXP_TABLE[i - 5] ^ EXP_TABLE[i - 6] ^ EXP_TABLE[i - 8];
  }
  for (let i = 0; i < 255; i++) LOG_TABLE[EXP_TABLE[i]] = i;
  return {
    glog(n) {
      if (n < 1) throw new Error("glog(" + n + ")");
      return LOG_TABLE[n];
    },
    gexp(n) {
      while (n < 0) n += 255;
      while (n >= 256) n -= 255;
      return EXP_TABLE[n];
    },
  };
})();

function qrPolynomial(num, shift) {
  let offset = 0;
  while (offset < num.length && num[offset] === 0) offset++;
  const _num = new Array(num.length - offset + shift);
  for (let i = 0; i < num.length - offset; i++) _num[i] = num[i + offset];
  return {
    getAt: (index) => _num[index],
    getLength: () => _num.length,
    multiply(e) {
      const out = new Array(_num.length + e.getLength() - 1).fill(0);
      for (let i = 0; i < _num.length; i++) {
        for (let j = 0; j < e.getLength(); j++) {
          out[i + j] ^= QRMath.gexp(QRMath.glog(_num[i]) + QRMath.glog(e.getAt(j)));
        }
      }
      return qrPolynomial(out, 0);
    },
    mod(e) {
      if (_num.length - e.getLength() < 0) return this;
      const ratio = QRMath.glog(_num[0]) - QRMath.glog(e.getAt(0));
      const out = _num.slice();
      for (let i = 0; i < e.getLength(); i++) {
        out[i] ^= QRMath.gexp(QRMath.glog(e.getAt(i)) + ratio);
      }
      return qrPolynomial(out, 0).mod(e);
    },
  };
}

const QRUtil = (function () {
  const PATTERN_POSITION_TABLE = [
    [], [6, 18], [6, 22], [6, 26], [6, 30], [6, 34],
    [6, 22, 38], [6, 24, 42], [6, 26, 46], [6, 28, 50], [6, 30, 54], [6, 32, 58], [6, 34, 62],
    [6, 26, 46, 66], [6, 26, 48, 70], [6, 26, 50, 74], [6, 30, 54, 78], [6, 30, 56, 82], [6, 30, 58, 86],
    [6, 34, 62, 90], [6, 28, 50, 72, 94], [6, 26, 50, 74, 98], [6, 30, 54, 78, 102], [6, 28, 54, 80, 106],
    [6, 32, 58, 84, 110], [6, 30, 58, 86, 114], [6, 34, 62, 90, 118], [6, 26, 50, 74, 98, 122],
    [6, 30, 54, 78, 102, 126], [6, 26, 52, 78, 104, 130], [6, 30, 56, 82, 108, 134], [6, 34, 60, 86, 112, 138],
    [6, 30, 58, 86, 114, 142], [6, 34, 62, 90, 118, 146], [6, 30, 54, 78, 102, 126, 150],
    [6, 24, 50, 76, 102, 128, 154], [6, 28, 54, 80, 106, 132, 158], [6, 32, 58, 84, 110, 136, 162],
    [6, 26, 54, 82, 110, 138, 166], [6, 30, 58, 86, 114, 142, 170],
  ];
  const G15 = (1 << 10) | (1 << 8) | (1 << 5) | (1 << 4) | (1 << 2) | (1 << 1) | (1 << 0);
  const G18 = (1 << 12) | (1 << 11) | (1 << 10) | (1 << 9) | (1 << 8) | (1 << 5) | (1 << 2) | (1 << 0);
  const G15_MASK = (1 << 14) | (1 << 12) | (1 << 10) | (1 << 4) | (1 << 1);

  function getBCHDigit(data) {
    let digit = 0;
    while (data !== 0) {
      digit++;
      data >>>= 1;
    }
    return digit;
  }

  return {
    getBCHTypeInfo(data) {
      let d = data << 10;
      while (getBCHDigit(d) - getBCHDigit(G15) >= 0) d ^= G15 << (getBCHDigit(d) - getBCHDigit(G15));
      return ((data << 10) | d) ^ G15_MASK;
    },
    getBCHTypeNumber(data) {
      let d = data << 12;
      while (getBCHDigit(d) - getBCHDigit(G18) >= 0) d ^= G18 << (getBCHDigit(d) - getBCHDigit(G18));
      return (data << 12) | d;
    },
    getPatternPosition: (typeNumber) => PATTERN_POSITION_TABLE[typeNumber - 1],
    getMaskFunction(maskPattern) {
      switch (maskPattern) {
        case QRMaskPattern.PATTERN000: return (i, j) => (i + j) % 2 === 0;
        case QRMaskPattern.PATTERN001: return (i) => i % 2 === 0;
        case QRMaskPattern.PATTERN010: return (i, j) => j % 3 === 0;
        case QRMaskPattern.PATTERN011: return (i, j) => (i + j) % 3 === 0;
        case QRMaskPattern.PATTERN100: return (i, j) => (Math.floor(i / 2) + Math.floor(j / 3)) % 2 === 0;
        case QRMaskPattern.PATTERN101: return (i, j) => ((i * j) % 2) + ((i * j) % 3) === 0;
        case QRMaskPattern.PATTERN110: return (i, j) => (((i * j) % 2) + ((i * j) % 3)) % 2 === 0;
        case QRMaskPattern.PATTERN111: return (i, j) => (((i * j) % 3) + ((i + j) % 2)) % 2 === 0;
        default: throw new Error("bad maskPattern:" + maskPattern);
      }
    },
    getErrorCorrectPolynomial(len) {
      let a = qrPolynomial([1], 0);
      for (let i = 0; i < len; i++) a = a.multiply(qrPolynomial([1, QRMath.gexp(i)], 0));
      return a;
    },
    getLengthInBits(mode, type) {
      if (type >= 1 && type < 10) return 8;
      if (type < 27) return 16;
      if (type < 41) return 16;
      throw new Error("type:" + type);
    },
    getLostPoint(qr) {
      const n = qr.getModuleCount();
      let lostPoint = 0;
      for (let row = 0; row < n; row++) {
        for (let col = 0; col < n; col++) {
          let sameCount = 0;
          const dark = qr.isDark(row, col);
          for (let r = -1; r <= 1; r++) {
            if (row + r < 0 || n <= row + r) continue;
            for (let c = -1; c <= 1; c++) {
              if (col + c < 0 || n <= col + c) continue;
              if (r === 0 && c === 0) continue;
              if (dark === qr.isDark(row + r, col + c)) sameCount++;
            }
          }
          if (sameCount > 5) lostPoint += 3 + sameCount - 5;
        }
      }
      for (let row = 0; row < n - 1; row++) {
        for (let col = 0; col < n - 1; col++) {
          let count = 0;
          if (qr.isDark(row, col)) count++;
          if (qr.isDark(row + 1, col)) count++;
          if (qr.isDark(row, col + 1)) count++;
          if (qr.isDark(row + 1, col + 1)) count++;
          if (count === 0 || count === 4) lostPoint += 3;
        }
      }
      for (let row = 0; row < n; row++) {
        for (let col = 0; col < n - 6; col++) {
          if (
            qr.isDark(row, col) && !qr.isDark(row, col + 1) && qr.isDark(row, col + 2) &&
            qr.isDark(row, col + 3) && qr.isDark(row, col + 4) && !qr.isDark(row, col + 5) && qr.isDark(row, col + 6)
          ) lostPoint += 40;
        }
      }
      for (let col = 0; col < n; col++) {
        for (let row = 0; row < n - 6; row++) {
          if (
            qr.isDark(row, col) && !qr.isDark(row + 1, col) && qr.isDark(row + 2, col) &&
            qr.isDark(row + 3, col) && qr.isDark(row + 4, col) && !qr.isDark(row + 5, col) && qr.isDark(row + 6, col)
          ) lostPoint += 40;
        }
      }
      let darkCount = 0;
      for (let col = 0; col < n; col++) for (let row = 0; row < n; row++) if (qr.isDark(row, col)) darkCount++;
      lostPoint += (Math.abs((100 * darkCount) / n / n - 50) / 5) * 10;
      return lostPoint;
    },
  };
})();

function qrBitBuffer() {
  const buffer = [];
  let length = 0;
  return {
    getBuffer: () => buffer,
    put(num, len) {
      for (let i = 0; i < len; i++) this.putBit(((num >>> (len - i - 1)) & 1) === 1);
    },
    getLengthInBits: () => length,
    putBit(bit) {
      const bufIndex = Math.floor(length / 8);
      if (buffer.length <= bufIndex) buffer.push(0);
      if (bit) buffer[bufIndex] |= 0x80 >>> length % 8;
      length++;
    },
  };
}

function qr8BitByte(data) {
  const bytes = [];
  for (let i = 0; i < data.length; i++) bytes.push(data.charCodeAt(i) & 0xff);
  return {
    getMode: () => QRMode.MODE_8BIT_BYTE,
    getLength: () => bytes.length,
    write(buffer) {
      for (let i = 0; i < bytes.length; i++) buffer.put(bytes[i], 8);
    },
  };
}

const QRRSBlock = (function () {
  const RS_BLOCK_TABLE = [
    [1, 26, 19], [1, 26, 16], [1, 26, 13], [1, 26, 9],
    [1, 44, 34], [1, 44, 28], [1, 44, 22], [1, 44, 16],
    [1, 70, 55], [1, 70, 44], [2, 35, 17], [2, 35, 13],
    [1, 100, 80], [2, 50, 32], [2, 50, 24], [4, 25, 9],
    [1, 134, 108], [2, 67, 43], [2, 33, 15, 2, 34, 16], [2, 33, 11, 2, 34, 12],
    [2, 86, 68], [4, 43, 27], [4, 43, 19], [4, 43, 15],
    [2, 98, 78], [4, 49, 31], [2, 32, 14, 4, 33, 15], [4, 39, 13, 1, 40, 14],
    [2, 121, 97], [2, 60, 38, 2, 61, 39], [4, 40, 18, 2, 41, 19], [4, 40, 14, 2, 41, 15],
    [2, 146, 116], [3, 58, 36, 2, 59, 37], [4, 36, 16, 4, 37, 17], [4, 36, 12, 4, 37, 13],
    [2, 86, 68, 2, 87, 69], [4, 69, 43, 1, 70, 44], [6, 43, 19, 2, 44, 20], [6, 43, 15, 2, 44, 16],
  ];
  const qrRSBlock = (totalCount, dataCount) => ({ totalCount, dataCount });
  function getRsBlockTable(typeNumber, level) {
    switch (level) {
      case QRErrorCorrectionLevel.L: return RS_BLOCK_TABLE[(typeNumber - 1) * 4 + 0];
      case QRErrorCorrectionLevel.M: return RS_BLOCK_TABLE[(typeNumber - 1) * 4 + 1];
      case QRErrorCorrectionLevel.Q: return RS_BLOCK_TABLE[(typeNumber - 1) * 4 + 2];
      case QRErrorCorrectionLevel.H: return RS_BLOCK_TABLE[(typeNumber - 1) * 4 + 3];
      default: return undefined;
    }
  }
  return {
    getRSBlocks(typeNumber, level) {
      const rsBlock = getRsBlockTable(typeNumber, level);
      if (!rsBlock) throw new Error("bad rs block @ typeNumber:" + typeNumber);
      const length = rsBlock.length / 3;
      const list = [];
      for (let i = 0; i < length; i++) {
        const count = rsBlock[i * 3 + 0];
        const totalCount = rsBlock[i * 3 + 1];
        const dataCount = rsBlock[i * 3 + 2];
        for (let j = 0; j < count; j++) list.push(qrRSBlock(totalCount, dataCount));
      }
      return list;
    },
  };
})();

function qrcodeCreate(typeNumber, errorCorrectionLevel) {
  const PAD0 = 0xec;
  const PAD1 = 0x11;
  const _errorCorrectionLevel = QRErrorCorrectionLevel[errorCorrectionLevel];
  let _modules = null;
  let _moduleCount = 0;
  let _dataCache = null;
  const _dataList = [];
  const _this = {};

  function setupPositionProbePattern(row, col) {
    for (let r = -1; r <= 7; r++) {
      if (row + r <= -1 || _moduleCount <= row + r) continue;
      for (let c = -1; c <= 7; c++) {
        if (col + c <= -1 || _moduleCount <= col + c) continue;
        if ((r >= 0 && r <= 6 && (c === 0 || c === 6)) || (c >= 0 && c <= 6 && (r === 0 || r === 6)) || (r >= 2 && r <= 4 && c >= 2 && c <= 4)) {
          _modules[row + r][col + c] = true;
        } else {
          _modules[row + r][col + c] = false;
        }
      }
    }
  }

  function setupTimingPattern() {
    for (let r = 8; r < _moduleCount - 8; r++) {
      if (_modules[r][6] != null) continue;
      _modules[r][6] = r % 2 === 0;
    }
    for (let c = 8; c < _moduleCount - 8; c++) {
      if (_modules[6][c] != null) continue;
      _modules[6][c] = c % 2 === 0;
    }
  }

  function setupPositionAdjustPattern() {
    const pos = QRUtil.getPatternPosition(typeNumber);
    for (let i = 0; i < pos.length; i++) {
      for (let j = 0; j < pos.length; j++) {
        const row = pos[i];
        const col = pos[j];
        if (_modules[row][col] != null) continue;
        for (let r = -2; r <= 2; r++) {
          for (let c = -2; c <= 2; c++) {
            if (r === -2 || r === 2 || c === -2 || c === 2 || (r === 0 && c === 0)) {
              _modules[row + r][col + c] = true;
            } else {
              _modules[row + r][col + c] = false;
            }
          }
        }
      }
    }
  }

  function setupTypeNumber(test) {
    const bits = QRUtil.getBCHTypeNumber(typeNumber);
    for (let i = 0; i < 18; i++) {
      const mod = !test && ((bits >> i) & 1) === 1;
      _modules[Math.floor(i / 3)][(i % 3) + _moduleCount - 8 - 3] = mod;
    }
    for (let i = 0; i < 18; i++) {
      const mod = !test && ((bits >> i) & 1) === 1;
      _modules[(i % 3) + _moduleCount - 8 - 3][Math.floor(i / 3)] = mod;
    }
  }

  function setupTypeInfo(test, maskPattern) {
    const data = (_errorCorrectionLevel << 3) | maskPattern;
    const bits = QRUtil.getBCHTypeInfo(data);
    for (let i = 0; i < 15; i++) {
      const mod = !test && ((bits >> i) & 1) === 1;
      if (i < 6) _modules[i][8] = mod;
      else if (i < 8) _modules[i + 1][8] = mod;
      else _modules[_moduleCount - 15 + i][8] = mod;
    }
    for (let i = 0; i < 15; i++) {
      const mod = !test && ((bits >> i) & 1) === 1;
      if (i < 8) _modules[8][_moduleCount - i - 1] = mod;
      else if (i < 9) _modules[8][15 - i - 1 + 1] = mod;
      else _modules[8][15 - i - 1] = mod;
    }
    _modules[_moduleCount - 8][8] = !test;
  }

  function mapData(data, maskPattern) {
    let inc = -1;
    let row = _moduleCount - 1;
    let bitIndex = 7;
    let byteIndex = 0;
    const maskFunc = QRUtil.getMaskFunction(maskPattern);
    for (let col = _moduleCount - 1; col > 0; col -= 2) {
      if (col === 6) col--;
      while (true) {
        for (let c = 0; c < 2; c++) {
          if (_modules[row][col - c] == null) {
            let dark = false;
            if (byteIndex < data.length) dark = ((data[byteIndex] >>> bitIndex) & 1) === 1;
            if (maskFunc(row, col - c)) dark = !dark;
            _modules[row][col - c] = dark;
            bitIndex--;
            if (bitIndex === -1) {
              byteIndex++;
              bitIndex = 7;
            }
          }
        }
        row += inc;
        if (row < 0 || _moduleCount <= row) {
          row -= inc;
          inc = -inc;
          break;
        }
      }
    }
  }

  function createBytes(buffer, rsBlocks) {
    let offset = 0;
    let maxDcCount = 0;
    let maxEcCount = 0;
    const dcdata = new Array(rsBlocks.length);
    const ecdata = new Array(rsBlocks.length);
    for (let r = 0; r < rsBlocks.length; r++) {
      const dcCount = rsBlocks[r].dataCount;
      const ecCount = rsBlocks[r].totalCount - dcCount;
      maxDcCount = Math.max(maxDcCount, dcCount);
      maxEcCount = Math.max(maxEcCount, ecCount);
      dcdata[r] = new Array(dcCount);
      for (let i = 0; i < dcdata[r].length; i++) dcdata[r][i] = 0xff & buffer.getBuffer()[i + offset];
      offset += dcCount;
      const rsPoly = QRUtil.getErrorCorrectPolynomial(ecCount);
      const rawPoly = qrPolynomial(dcdata[r], rsPoly.getLength() - 1);
      const modPoly = rawPoly.mod(rsPoly);
      ecdata[r] = new Array(rsPoly.getLength() - 1);
      for (let i = 0; i < ecdata[r].length; i++) {
        const modIndex = i + modPoly.getLength() - ecdata[r].length;
        ecdata[r][i] = modIndex >= 0 ? modPoly.getAt(modIndex) : 0;
      }
    }
    let totalCodeCount = 0;
    for (let i = 0; i < rsBlocks.length; i++) totalCodeCount += rsBlocks[i].totalCount;
    const data = new Array(totalCodeCount);
    let index = 0;
    for (let i = 0; i < maxDcCount; i++) {
      for (let r = 0; r < rsBlocks.length; r++) {
        if (i < dcdata[r].length) data[index++] = dcdata[r][i];
      }
    }
    for (let i = 0; i < maxEcCount; i++) {
      for (let r = 0; r < rsBlocks.length; r++) {
        if (i < ecdata[r].length) data[index++] = ecdata[r][i];
      }
    }
    return data;
  }

  function createData(typeNum, level, dataList) {
    const rsBlocks = QRRSBlock.getRSBlocks(typeNum, level);
    const buffer = qrBitBuffer();
    for (let i = 0; i < dataList.length; i++) {
      const data = dataList[i];
      buffer.put(data.getMode(), 4);
      buffer.put(data.getLength(), QRUtil.getLengthInBits(data.getMode(), typeNum));
      data.write(buffer);
    }
    let totalDataCount = 0;
    for (let i = 0; i < rsBlocks.length; i++) totalDataCount += rsBlocks[i].dataCount;
    if (buffer.getLengthInBits() > totalDataCount * 8) {
      throw new Error("code length overflow. (" + buffer.getLengthInBits() + ">" + totalDataCount * 8 + ")");
    }
    if (buffer.getLengthInBits() + 4 <= totalDataCount * 8) buffer.put(0, 4);
    while (buffer.getLengthInBits() % 8 !== 0) buffer.putBit(false);
    while (true) {
      if (buffer.getLengthInBits() >= totalDataCount * 8) break;
      buffer.put(PAD0, 8);
      if (buffer.getLengthInBits() >= totalDataCount * 8) break;
      buffer.put(PAD1, 8);
    }
    return createBytes(buffer, rsBlocks);
  }

  function makeImpl(test, maskPattern) {
    _moduleCount = typeNumber * 4 + 17;
    _modules = new Array(_moduleCount);
    for (let row = 0; row < _moduleCount; row++) {
      _modules[row] = new Array(_moduleCount).fill(null);
    }
    setupPositionProbePattern(0, 0);
    setupPositionProbePattern(_moduleCount - 7, 0);
    setupPositionProbePattern(0, _moduleCount - 7);
    setupPositionAdjustPattern();
    setupTimingPattern();
    setupTypeInfo(test, maskPattern);
    if (typeNumber >= 7) setupTypeNumber(test);
    if (_dataCache == null) _dataCache = createData(typeNumber, _errorCorrectionLevel, _dataList);
    mapData(_dataCache, maskPattern);
  }

  function getBestMaskPattern() {
    let minLostPoint = 0;
    let pattern = 0;
    for (let i = 0; i < 8; i++) {
      makeImpl(true, i);
      const lostPoint = QRUtil.getLostPoint(_this);
      if (i === 0 || minLostPoint > lostPoint) {
        minLostPoint = lostPoint;
        pattern = i;
      }
    }
    return pattern;
  }

  _this.addData = (data) => {
    _dataList.push(qr8BitByte(data));
    _dataCache = null;
  };
  _this.isDark = (row, col) => {
    if (row < 0 || _moduleCount <= row || col < 0 || _moduleCount <= col) throw new Error(row + "," + col);
    return _modules[row][col];
  };
  _this.getModuleCount = () => _moduleCount;
  _this.make = () => makeImpl(false, getBestMaskPattern());

  return _this;
}

function makeQRCode(text) {
  for (let typeNumber = 1; typeNumber <= 10; typeNumber++) {
    try {
      const qr = qrcodeCreate(typeNumber, "L");
      qr.addData(text);
      qr.make();
      return qr;
    } catch (e) {
      // too small for this typeNumber — try the next one
    }
  }
  return null;
}

function QRCodeSVG({ value, size = 176 }) {
  const qr = React.useMemo(() => {
    try {
      return makeQRCode(value);
    } catch (e) {
      return null;
    }
  }, [value]);
  if (!qr) return null;
  const count = qr.getModuleCount();
  const cell = size / count;
  const rects = [];
  for (let r = 0; r < count; r++) {
    for (let c = 0; c < count; c++) {
      if (qr.isDark(r, c)) {
        rects.push(<rect key={`${r}-${c}`} x={c * cell} y={r * cell} width={cell + 0.5} height={cell + 0.5} />);
      }
    }
  }
  return (
    <svg
      width={size}
      height={size}
      viewBox={`0 0 ${size} ${size}`}
      shapeRendering="crispEdges"
      style={{ background: "#ffffff", borderRadius: 10, padding: 10, boxSizing: "content-box" }}
    >
      <g fill="#000000">{rects}</g>
    </svg>
  );
}

function newJudgeId(judges) {
  let n = judges.length + 1;
  while (judges.some((j) => j.id === `j${n}`)) n++;
  return `j${n}`;
}

const DEFAULT_CONFIG = {
  adminPassword: "admin123",
  appUrl: "",
  locked: false,
  presentationCriteria: [
    { name: "İçerik", weight: 1, description: "" },
    { name: "Sunum Becerisi", weight: 1, description: "" },
    { name: "Yaratıcılık", weight: 1, description: "" },
  ],
  presentationGroups: [{ id: "pg1", name: "Sunum Grubu 1" }],
  judges: Array.from({ length: 5 }, (_, i) => ({
    id: `j${i + 1}`,
    name: `Hakem ${i + 1}`,
    username: `hakem${i + 1}`,
    password: "1234",
  })),
  groups: [{ id: "g1", name: "Grup 1" }],
  people: Array.from({ length: 10 }, (_, i) => ({
    name: `Kişi ${i + 1}`,
    groupId: "g1",
    photo: "",
    school: "",
    department: "",
    featured: false,
  })),
  criteria: [
    { name: "Kriter 1", weight: 1, description: "" },
    { name: "Kriter 2", weight: 1, description: "" },
    { name: "Kriter 3", weight: 1, description: "" },
  ],
};

function normalizeConfig(raw) {
  const cfg = { ...DEFAULT_CONFIG, ...raw };
  if (!cfg.groups || cfg.groups.length === 0) {
    cfg.groups = [{ id: "g1", name: "Grup 1" }];
  }
  cfg.criteria = (cfg.criteria || []).map((c, i) => {
    if (typeof c === "string") return { name: c, weight: 1, description: "" };
    return { name: c.name ?? `Kriter ${i + 1}`, weight: typeof c.weight === "number" ? c.weight : 1, description: c.description ?? "" };
  });
  cfg.people = (cfg.people || []).map((p, i) => {
    if (typeof p === "string") return { name: p, groupId: cfg.groups[0].id, photo: "", school: "", department: "", featured: false };
    return {
      name: p.name ?? `Kişi ${i + 1}`,
      groupId: p.groupId ?? cfg.groups[0].id,
      photo: p.photo ?? "",
      school: p.school ?? "",
      department: p.department ?? "",
      featured: !!p.featured,
    };
  });
  cfg.judges = (cfg.judges || []).map((j, i) => {
    if (typeof j === "string") {
      return { id: `j${i + 1}`, name: j, username: `hakem${i + 1}`, password: "1234" };
    }
    return {
      id: j.id ?? `j${i + 1}`,
      name: j.name ?? `Hakem ${i + 1}`,
      username: j.username ?? `hakem${i + 1}`,
      password: j.password ?? "1234",
    };
  });
  cfg.adminPassword = cfg.adminPassword ?? DEFAULT_CONFIG.adminPassword;
  cfg.appUrl = cfg.appUrl ?? "";
  cfg.locked = !!cfg.locked;
  cfg.presentationCriteria = (cfg.presentationCriteria || []).map((c, i) => {
    if (typeof c === "string") return { name: c, weight: 1, description: "" };
    return { name: c.name ?? `Sunum Kriteri ${i + 1}`, weight: typeof c.weight === "number" ? c.weight : 1, description: c.description ?? "" };
  });
  if (!cfg.presentationGroups || cfg.presentationGroups.length === 0) {
    cfg.presentationGroups = [{ id: "pg1", name: "Sunum Grubu 1" }];
  }
  return cfg;
}

function resizeImageFile(file, maxDim = 220, quality = 0.75) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Dosya okunamadı"));
    reader.onload = () => {
      const img = new Image();
      img.onerror = () => reject(new Error("Görsel yüklenemedi"));
      img.onload = () => {
        let { width, height } = img;
        if (width > height && width > maxDim) {
          height = Math.round((height * maxDim) / width);
          width = maxDim;
        } else if (height > maxDim) {
          width = Math.round((width * maxDim) / height);
          height = maxDim;
        }
        const canvas = document.createElement("canvas");
        canvas.width = width;
        canvas.height = height;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0, width, height);
        resolve(canvas.toDataURL("image/jpeg", quality));
      };
      img.src = reader.result;
    };
    reader.readAsDataURL(file);
  });
}

function newGroupId(groups) {
  let n = groups.length + 1;
  while (groups.some((g) => g.id === `g${n}`)) n++;
  return `g${n}`;
}

function newPresGroupId(groups) {
  let n = groups.length + 1;
  while (groups.some((g) => g.id === `pg${n}`)) n++;
  return `pg${n}`;
}

// Retries a storage.set once after a short delay if the first attempt fails
// or throws. With 50 people hitting the backend at once, a transient
// timeout is more likely than in single-user testing — one retry meaningfully
// reduces "kayıt başarısız" complaints without masking a real, persistent failure.
async function setWithRetry(key, value, delayMs = 900) {
  try {
    const result = await storage.set(key, value, true);
    if (result) return result;
  } catch (e) {}
  await new Promise((resolve) => setTimeout(resolve, delayMs));
  try {
    return await storage.set(key, value, true);
  } catch (e) {
    return null;
  }
}

function Avatar({ photo, name, size = 40, shape = "circle" }) {
  const style = {
    width: size,
    height: size,
    fontSize: Math.round(size * 0.4),
    borderRadius: shape === "square" ? Math.round(size * 0.18) : "50%",
  };
  if (photo) {
    return <img src={photo} alt={name} className="sb-avatar" style={style} />;
  }
  const initial = (name || "?").trim().charAt(0).toUpperCase();
  return (
    <div className="sb-avatar sb-avatar-fallback" style={style}>
      {initial}
    </div>
  );
}

function ScoreDigits({ value, size = "lg" }) {
  const text = value === null || value === undefined ? "—" : value.toFixed(1);
  return (
    <div className={`flip-group flip-${size}`}>
      {text.split("").map((ch, i) => (
        <span key={i} className={ch === "." ? "flip-dot" : "flip-digit"}>
          {ch}
        </span>
      ))}
    </div>
  );
}

function ScoreBlock({ items, view }) {
  if (items.length === 0) return <div className="sb-empty">Bu grupta yarışmacı yok.</div>;
  if (view === "list") {
    return (
      <>
        {items.map((p, i) => (
          <div className="sb-rankrow" key={p.idx}>
            <div className="sb-rankpos">{i + 1}</div>
            <Avatar photo={p.photo} name={p.name} size={32} />
            <div style={{ flex: 1 }}>
              <div className="sb-rankname">{p.name}</div>
              <div className="sb-rankmeta">
                {p.count}/{p.total} hakem oy verdi
              </div>
              {p.notes && p.notes.length > 0 && (
                <div className="sb-rownotes">
                  {p.notes.map((n, ni) => (
                    <div className="sb-rownote" key={ni}>
                      <b>{n.judgeName}:</b> {n.note}
                    </div>
                  ))}
                </div>
              )}
            </div>
            <ScoreDigits value={p.overall} size="sm" />
          </div>
        ))}
      </>
    );
  }
  if (items.every((p) => p.overall === null)) {
    return <div className="sb-empty">Henüz grafik için yeterli puan yok.</div>;
  }

  const renderYAxisTick = (props) => {
    const { x, y, index } = props;
    const item = items[index];
    if (!item) return null;
    return (
      <g transform={`translate(${x},${y})`}>
        <foreignObject x={-142} y={-15} width={134} height={30}>
          <div
            xmlns="http://www.w3.org/1999/xhtml"
            style={{ display: "flex", alignItems: "center", gap: 7, height: "100%" }}
          >
            {item.photo ? (
              <img
                src={item.photo}
                style={{ width: 24, height: 24, borderRadius: 7, objectFit: "cover", flexShrink: 0 }}
              />
            ) : (
              <div
                style={{
                  width: 24, height: 24, borderRadius: 7, background: "#e2e4e9", color: "#6b7280",
                  fontSize: 10, fontWeight: 700, display: "flex", alignItems: "center",
                  justifyContent: "center", flexShrink: 0, fontFamily: "Oswald, sans-serif",
                }}
              >
                {(item.name || "?").trim().charAt(0).toUpperCase()}
              </div>
            )}
            <span
              style={{
                fontSize: 12, color: "#14171c", fontFamily: "Inter, sans-serif",
                whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
              }}
            >
              {item.name}
            </span>
          </div>
        </foreignObject>
      </g>
    );
  };

  const renderScoreLabel = (props) => {
    const { x, y, width, height, value } = props;
    if (value === null || value === undefined || value === 0) return null;
    const label = Number(value).toFixed(1);
    const boxW = 32;
    return (
      <g transform={`translate(${x + width + 7}, ${y + height / 2 - 9.5})`}>
        <rect width={boxW} height={19} rx={6} fill="#14171c" stroke="#14171c" />
        <text
          x={boxW / 2}
          y={13.5}
          textAnchor="middle"
          fontSize={11}
          fontFamily="JetBrains Mono, monospace"
          fontWeight={700}
          fill="#ffffff"
        >
          {label}
        </text>
      </g>
    );
  };

  return (
    <ResponsiveContainer width="100%" height={Math.max(160, items.length * 44)}>
      <BarChart
        data={items.map((p) => ({ name: p.name, score: p.overall ?? 0, count: p.count, total: p.total }))}
        layout="vertical"
        barCategoryGap="30%"
        margin={{ top: 4, right: 40, left: 8, bottom: 4 }}
      >
        <defs>
          <linearGradient id="sbBarGold" x1="0" y1="0" x2="1" y2="0">
            <stop offset="0%" stopColor="#c99a1e" />
            <stop offset="100%" stopColor="#f2c230" />
          </linearGradient>
          <linearGradient id="sbBarRed" x1="0" y1="0" x2="1" y2="0">
            <stop offset="0%" stopColor="#9c1f29" />
            <stop offset="100%" stopColor="#e63946" />
          </linearGradient>
        </defs>
        <CartesianGrid strokeDasharray="3 3" stroke="#e2e4e9" horizontal={false} />
        <XAxis
          type="number"
          domain={[0, 5]}
          tick={{ fill: "#6b7280", fontSize: 11, fontFamily: "JetBrains Mono, monospace" }}
          axisLine={{ stroke: "#d1d5db" }}
          tickLine={false}
        />
        <YAxis type="category" dataKey="name" width={142} tick={renderYAxisTick} axisLine={{ stroke: "#d1d5db" }} tickLine={false} />
        <Tooltip
          cursor={{ fill: "rgba(234,179,8,0.08)" }}
          contentStyle={{
            background: "#14171c",
            border: "1px solid #14171c",
            borderRadius: 8,
            fontFamily: "Inter, sans-serif",
            fontSize: 12.5,
          }}
          labelStyle={{ color: "#ffffff", fontWeight: 600 }}
          formatter={(value, name, props) => [
            `${Number(value).toFixed(1)} · ${props.payload.count}/${props.payload.total} hakem`,
            "Ortalama",
          ]}
        />
        <Bar dataKey="score" radius={[0, 8, 8, 0]} maxBarSize={26} isAnimationActive>
          {items.map((p, i) => (
            <Cell
              key={p.idx}
              fill={i === 0 ? "url(#sbBarGold)" : "url(#sbBarRed)"}
              fillOpacity={p.overall === null ? 0.2 : 1}
            />
          ))}
          <LabelList dataKey="score" content={renderScoreLabel} />
        </Bar>
      </BarChart>
    </ResponsiveContainer>
  );
}

function FireworksCanvas({ active }) {
  const canvasRef = useRef(null);

  useEffect(() => {
    if (!active) return;
    const canvas = canvasRef.current;
    if (!canvas) return;
    const ctx = canvas.getContext("2d");
    let width = (canvas.width = canvas.offsetWidth);
    let height = (canvas.height = canvas.offsetHeight);
    const handleResize = () => {
      width = canvas.width = canvas.offsetWidth;
      height = canvas.height = canvas.offsetHeight;
    };
    window.addEventListener("resize", handleResize);

    const colors = ["#eab308", "#e63946", "#14171c"];
    let particles = [];

    function spawnBurst(x, y) {
      const count = 46;
      for (let i = 0; i < count; i++) {
        const angle = (Math.PI * 2 * i) / count + Math.random() * 0.25;
        const speed = 2.2 + Math.random() * 3.2;
        particles.push({
          x, y,
          vx: Math.cos(angle) * speed,
          vy: Math.sin(angle) * speed,
          life: 1,
          color: colors[Math.floor(Math.random() * colors.length)],
          size: 2 + Math.random() * 2.2,
        });
      }
    }

    const timers = [];
    [0, 280, 560, 900, 1300, 1750].forEach((delay) => {
      timers.push(
        setTimeout(() => {
          spawnBurst(width * (0.18 + Math.random() * 0.64), height * (0.16 + Math.random() * 0.38));
        }, delay)
      );
    });

    let running = true;
    let raf;
    function tick() {
      if (!running) return;
      ctx.clearRect(0, 0, width, height);
      particles.forEach((p) => {
        p.x += p.vx;
        p.y += p.vy;
        p.vy += 0.045;
        p.life -= 0.011;
      });
      particles = particles.filter((p) => p.life > 0);
      particles.forEach((p) => {
        ctx.globalAlpha = Math.max(p.life, 0);
        ctx.fillStyle = p.color;
        ctx.beginPath();
        ctx.arc(p.x, p.y, p.size, 0, Math.PI * 2);
        ctx.fill();
      });
      ctx.globalAlpha = 1;
      raf = requestAnimationFrame(tick);
    }
    tick();

    const stopTimer = setTimeout(() => {
      running = false;
      if (raf) cancelAnimationFrame(raf);
      ctx.clearRect(0, 0, width, height);
    }, 4500);

    return () => {
      running = false;
      if (raf) cancelAnimationFrame(raf);
      timers.forEach(clearTimeout);
      clearTimeout(stopTimer);
      window.removeEventListener("resize", handleResize);
    };
  }, [active]);

  if (!active) return null;
  return <canvas ref={canvasRef} className="sb-fireworks-canvas" />;
}

const COMPARE_COLORS = ["#eab308", "#e63946", "#14171c", "#3b82f6"];

function CompareView({ people, criteria, personCriteriaAverages, selection, setSelection }) {
  const toggle = (idx) => {
    setSelection((prev) => {
      if (prev.includes(idx)) return prev.filter((i) => i !== idx);
      if (prev.length >= 4) return prev;
      return [...prev, idx];
    });
  };

  const selectedPeople = selection.map((idx) => people.find((p) => p.idx === idx)).filter(Boolean);

  const radarData = criteria.map((c) => {
    const row = { criterion: c.name };
    selectedPeople.forEach((p) => {
      const avgs = personCriteriaAverages(p.idx);
      row[p.name] = avgs[c.name] === "" ? 0 : avgs[c.name];
    });
    return row;
  });

  return (
    <>
      <div className="sb-card">
        <div className="sb-label">Karşılaştırılacak kişileri seç (en fazla 4)</div>
        <div className="sb-peoplegrid" style={{ gridTemplateColumns: "repeat(auto-fill, minmax(130px, 1fr))" }}>
          {people.map((p) => {
            const active = selection.includes(p.idx);
            const disabled = !active && selection.length >= 4;
            return (
              <button
                key={p.idx}
                type="button"
                className={`sb-personcard ${active ? "active" : ""}`}
                style={{ aspectRatio: "auto", padding: "10px 8px", opacity: disabled ? 0.4 : 1 }}
                onClick={() => !disabled && toggle(p.idx)}
                disabled={disabled}
              >
                <div className="ptop">
                  <Avatar photo={p.photo} name={p.name} size={44} shape="square" />
                  <div className="pname">{p.name}</div>
                </div>
              </button>
            );
          })}
        </div>
      </div>

      <div className="sb-card">
        {selectedPeople.length === 0 ? (
          <div className="sb-empty">Karşılaştırmak için en az bir kişi seç.</div>
        ) : (
          <>
            <ResponsiveContainer width="100%" height={340}>
              <RadarChart data={radarData} outerRadius="70%">
                <PolarGrid stroke="#e2e4e9" />
                <PolarAngleAxis dataKey="criterion" tick={{ fill: "#14171c", fontSize: 12, fontFamily: "Inter, sans-serif" }} />
                <PolarRadiusAxis domain={[0, 5]} tick={{ fill: "#6b7280", fontSize: 10 }} stroke="#e2e4e9" />
                {selectedPeople.map((p, i) => (
                  <Radar
                    key={p.idx}
                    name={p.name}
                    dataKey={p.name}
                    stroke={COMPARE_COLORS[i % COMPARE_COLORS.length]}
                    fill={COMPARE_COLORS[i % COMPARE_COLORS.length]}
                    fillOpacity={0.15}
                    strokeWidth={2}
                  />
                ))}
                <Legend wrapperStyle={{ fontSize: 12, fontFamily: "Inter, sans-serif", color: "#14171c" }} />
                <Tooltip
                  contentStyle={{ background: "#14171c", border: "1px solid #14171c", borderRadius: 8, fontSize: 12.5 }}
                  labelStyle={{ color: "#ffffff", fontWeight: 600 }}
                />
              </RadarChart>
            </ResponsiveContainer>
          </>
        )}
      </div>
    </>
  );
}

function ReportView({ config, rankedPeople, personCriteriaAverages, completionPct, totalDoneRatings, totalPossibleRatings, onClose }) {
  const dateStr = new Date().toLocaleDateString("tr-TR", { day: "numeric", month: "long", year: "numeric" });
  return (
    <div className="sb-reportpage">
      <style>{`
        .sb-reportpage {
          background: #ffffff;
          color: #14171c;
          min-height: 100vh;
          font-family: 'Inter', sans-serif;
          padding: 24px;
        }
        .sb-report-toolbar {
          display: flex;
          gap: 10px;
          justify-content: center;
          margin-bottom: 24px;
        }
        .sb-report-toolbar .sb-savebtn {
          background: #14171c;
          color: #ffffff;
          border-radius: 8px;
          font-weight: 600;
          font-size: 13px;
          cursor: pointer;
          border: none;
          display: inline-flex;
          align-items: center;
          gap: 8px;
        }
        .sb-copybtn2 {
          padding: 10px 18px;
          border-radius: 8px;
          border: 1px solid #cfcfcf;
          background: #ffffff;
          color: #14171c;
          font-weight: 600;
          font-size: 13px;
          cursor: pointer;
        }
        .sb-report {
          max-width: 900px;
          margin: 0 auto;
          border: 1px solid #e2e2e2;
          border-radius: 12px;
          padding: 36px 40px;
        }
        .sb-report-head {
          display: flex;
          justify-content: space-between;
          align-items: baseline;
          border-bottom: 3px solid #14171c;
          padding-bottom: 16px;
          margin-bottom: 24px;
        }
        .sb-report-eyebrow {
          font-family: 'Oswald', sans-serif;
          font-weight: 700;
          font-size: 24px;
          text-transform: uppercase;
          letter-spacing: 0.02em;
        }
        .sb-report-date { font-size: 13px; color: #666; }
        .sb-report-stats {
          display: grid;
          grid-template-columns: repeat(4, 1fr);
          gap: 14px;
          margin-bottom: 30px;
        }
        .sb-report-stat {
          border: 1px solid #e2e2e2;
          border-radius: 10px;
          padding: 14px;
          text-align: center;
        }
        .sb-report-statval {
          font-family: 'Oswald', sans-serif;
          font-weight: 700;
          font-size: 26px;
          color: #e63946;
        }
        .sb-report-statlabel { font-size: 11px; color: #666; margin-top: 4px; }
        .sb-report-table { width: 100%; border-collapse: collapse; font-size: 12.5px; }
        .sb-report-table th {
          text-align: left;
          border-bottom: 2px solid #14171c;
          padding: 8px 6px;
          font-size: 11px;
          text-transform: uppercase;
          letter-spacing: 0.03em;
          color: #444;
        }
        .sb-report-table td {
          padding: 8px 6px;
          border-bottom: 1px solid #eee;
        }
        .sb-report-top { background: #fff8e6; font-weight: 600; }
        .sb-report-overall { font-weight: 700; color: #e63946; }
        .sb-report-footer {
          margin-top: 26px;
          font-size: 10.5px;
          color: #999;
          text-align: center;
        }
        @media print {
          .sb-noprint { display: none !important; }
          .sb-reportpage { padding: 0; }
          .sb-report { border: none; max-width: 100%; }
        }
      `}</style>
      <div className="sb-report-toolbar sb-noprint">
        <button className="sb-savebtn" style={{ width: "auto", padding: "10px 18px" }} onClick={() => window.print()}>
          Yazdır / PDF Olarak Kaydet
        </button>
        <button className="sb-copybtn2" onClick={onClose}>
          Kapat
        </button>
      </div>
      <div className="sb-report">
        <div className="sb-report-head">
          <div className="sb-report-eyebrow">PUANLAMA SONUÇ RAPORU</div>
          <div className="sb-report-date">{dateStr}</div>
        </div>
        <div className="sb-report-stats">
          <div className="sb-report-stat">
            <div className="sb-report-statval">{config.people.length}</div>
            <div className="sb-report-statlabel">Yarışmacı</div>
          </div>
          <div className="sb-report-stat">
            <div className="sb-report-statval">{config.judges.length}</div>
            <div className="sb-report-statlabel">Hakem</div>
          </div>
          <div className="sb-report-stat">
            <div className="sb-report-statval">{config.criteria.length}</div>
            <div className="sb-report-statlabel">Kriter</div>
          </div>
          <div className="sb-report-stat">
            <div className="sb-report-statval">%{completionPct}</div>
            <div className="sb-report-statlabel">
              Tamamlanma ({totalDoneRatings}/{totalPossibleRatings})
            </div>
          </div>
        </div>

        <table className="sb-report-table">
          <thead>
            <tr>
              <th>#</th>
              <th>Yarışmacı</th>
              <th>Grup</th>
              {config.criteria.map((c) => (
                <th key={c.name}>{c.name}</th>
              ))}
              <th>Ortalama</th>
              <th>Hakem</th>
            </tr>
          </thead>
          <tbody>
            {rankedPeople.map((p, i) => {
              const avgs = personCriteriaAverages(p.idx);
              const groupName = config.groups.find((g) => g.id === p.groupId)?.name ?? "";
              return (
                <tr key={p.idx} className={i < 3 ? "sb-report-top" : ""}>
                  <td>{i + 1}</td>
                  <td>{p.name}</td>
                  <td>{groupName}</td>
                  {config.criteria.map((c) => (
                    <td key={c.name}>{avgs[c.name] === "" ? "—" : avgs[c.name]}</td>
                  ))}
                  <td className="sb-report-overall">{p.overall !== null ? p.overall.toFixed(2) : "—"}</td>
                  <td>
                    {p.count}/{p.total}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>

        <div className="sb-report-footer">Bu rapor canlı puanlama sistemi üzerinden otomatik oluşturulmuştur.</div>
      </div>
    </div>
  );
}

export default function ScoreboardApp() {
  const [session, setSession] = useState(null); // null | {role:'judge', judgeId} | {role:'admin'}
  const [loginRole, setLoginRole] = useState("judge");
  const [loginUsername, setLoginUsername] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginAdminPassword, setLoginAdminPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [configLoaded, setConfigLoaded] = useState(false);
  const [timer, setTimer] = useState({ durationSec: 300, endAt: null, running: false });
  const [now, setNow] = useState(Date.now());
  const [pageUrl, setPageUrl] = useState("");

  useEffect(() => {
    try {
      setPageUrl(window.location.href);
    } catch (e) {}
  }, []);

  const [tab, setTab] = useState("score");
  const [dashView, setDashView] = useState("chart");
  const [presDashView, setPresDashView] = useState("list");
  const [dashGroupFilter, setDashGroupFilter] = useState("all");
  const [compareSelection, setCompareSelection] = useState([]);
  const [showReport, setShowReport] = useState(false);
  const [scoreGroupFilter, setScoreGroupFilter] = useState("all");
  const [config, setConfig] = useState(DEFAULT_CONFIG);
  const [ratings, setRatings] = useState({});
  const [adminActingJudgeId, setAdminActingJudgeId] = useState(null);
  const [selectedPerson, setSelectedPerson] = useState(null);
  const [draftScores, setDraftScores] = useState({});
  const [draftNote, setDraftNote] = useState("");
  const [showOnlyUnscored, setShowOnlyUnscored] = useState(false);
  const [presRatings, setPresRatings] = useState({});
  const [selectedPresGroup, setSelectedPresGroup] = useState(null);
  const [draftPresScores, setDraftPresScores] = useState({});
  const [draftPresNote, setDraftPresNote] = useState("");
  const [presSaving, setPresSaving] = useState(false);
  const [presSavedFlash, setPresSavedFlash] = useState(false);
  const [presSaveError, setPresSaveError] = useState("");
  const [saving, setSaving] = useState(false);
  const [savedFlash, setSavedFlash] = useState(false);
  const [saveError, setSaveError] = useState("");
  const [exportError, setExportError] = useState("");
  const [draftConfig, setDraftConfig] = useState(DEFAULT_CONFIG);
  const ratingsRef = useRef({});
  const configRef = useRef(DEFAULT_CONFIG);

  useEffect(() => {
    ratingsRef.current = ratings;
  }, [ratings]);
  useEffect(() => {
    configRef.current = config;
  }, [config]);

  const loadConfig = useCallback(async () => {
    try {
      const res = await storage.get("config", true);
      if (res && res.value) {
        const parsed = normalizeConfig(JSON.parse(res.value));
        setConfig(parsed);
        setConfigLoaded(true);
        return parsed;
      }
      // res resolved but carried no value — the key genuinely doesn't exist
      // yet, so it's safe to seed it. This only happens once, ever.
      await storage.set("config", JSON.stringify(DEFAULT_CONFIG), true);
      setConfigLoaded(true);
      return DEFAULT_CONFIG;
    } catch (e) {
      // The GET itself failed (network blip, parse error, etc). We must NOT
      // write DEFAULT_CONFIG here — that would silently wipe a real, already
      //-configured event if this poll (which now runs every 15s, not just
      // once at mount) happens to hit a transient error mid-event. Just keep
      // whatever config is already in local state and try again next poll.
      setConfigLoaded(true);
      return null;
    }
  }, []);

  useEffect(() => {
    (async () => {
      const initial = await loadConfig();
      // draftConfig only needs the initial snapshot — the admin's in-progress
      // edits shouldn't be clobbered by the periodic refresh below. Fall back
      // to the default shape if the very first load hit a transient error,
      // so the settings form never receives null.
      setDraftConfig(initial || DEFAULT_CONFIG);
    })();
    const id = setInterval(loadConfig, 15000);
    return () => clearInterval(id);
    // eslint-disable-next-line
  }, [loadConfig]);

  const loadAllRatings = useCallback(async () => {
    try {
      // Single combined bulk read for both person ratings (r:) and
      // presentation ratings (pr:) — one HTTP round trip instead of two.
      // At 50 concurrent users this halves the polling load on the backend.
      if (storage && typeof storage.bulk === "function") {
        const entries = await storage.bulk("r:,pr:");
        const parsedR = {};
        const parsedPR = {};
        Object.entries(entries).forEach(([k, v]) => {
          let val;
          try { val = JSON.parse(v); } catch (e) { return; }
          if (k.startsWith("pr:")) parsedPR[k] = val;
          else if (k.startsWith("r:")) parsedR[k] = val;
        });
        setRatings(parsedR);
        setPresRatings(parsedPR);
      } else {
        // Fallback for environments without a bulk endpoint (e.g. testing
        // inside the Claude artifact sandbox): two sequential list+get passes.
        const loadPrefix = async (prefix) => {
          const parsed = {};
          const list = await storage.list(prefix, true);
          if (list && list.keys && list.keys.length > 0) {
            const results = await Promise.all(
              list.keys.map(async (k) => {
                try {
                  const r = await storage.get(k, true);
                  return [k, r ? JSON.parse(r.value) : null];
                } catch (e) { return [k, null]; }
              })
            );
            results.forEach(([k, v]) => { if (v) parsed[k] = v; });
          }
          return parsed;
        };
        const [parsedR, parsedPR] = await Promise.all([loadPrefix("r:"), loadPrefix("pr:")]);
        setRatings(parsedR);
        setPresRatings(parsedPR);
      }
    } catch (e) {
      setRatings({});
      setPresRatings({});
    }
  }, []);

  useEffect(() => {
    loadAllRatings();
    const id = setInterval(loadAllRatings, 12000);
    return () => clearInterval(id);
  }, [loadAllRatings]);

  const loadTimer = useCallback(async () => {
    try {
      const res = await storage.get("timer", true);
      if (res && res.value) setTimer(JSON.parse(res.value));
    } catch (e) {}
  }, []);

  useEffect(() => {
    loadTimer();
    const id = setInterval(loadTimer, 8000);
    return () => clearInterval(id);
  }, [loadTimer]);

  useEffect(() => {
    const id = setInterval(() => setNow(Date.now()), 1000);
    return () => clearInterval(id);
  }, []);

  const saveTimer = async (next) => {
    setTimer(next);
    try {
      await storage.set("timer", JSON.stringify(next), true);
    } catch (e) {}
  };

  const timerRemainingSec = timer.running
    ? Math.max(0, Math.round((timer.endAt - now) / 1000))
    : timer.durationSec;

  const startTimer = () => saveTimer({ ...timer, endAt: Date.now() + timerRemainingSec * 1000, running: true });
  const pauseTimer = () => saveTimer({ ...timer, durationSec: timerRemainingSec, endAt: null, running: false });
  const resetTimer = (sec) => saveTimer({ durationSec: sec ?? timer.durationSec, endAt: null, running: false });

  const [judgeLocks, setJudgeLocks] = useState({});

  const loadJudgeLocks = useCallback(async () => {
    try {
      let parsed = {};
      if (storage && typeof storage.bulk === "function") {
        const entries = await storage.bulk("jl:");
        Object.entries(entries).forEach(([k, v]) => {
          const judgeId = k.slice(3); // strip "jl:" prefix
          try { parsed[judgeId] = JSON.parse(v); } catch (e) {}
        });
      } else {
        const list = await storage.list("jl:", true);
        if (list && list.keys && list.keys.length > 0) {
          const results = await Promise.all(
            list.keys.map(async (k) => {
              try {
                const r = await storage.get(k, true);
                return [k.slice(3), r ? JSON.parse(r.value) : null];
              } catch (e) { return [k.slice(3), null]; }
            })
          );
          results.forEach(([judgeId, v]) => { if (v !== null) parsed[judgeId] = v; });
        }
      }
      setJudgeLocks(parsed);
    } catch (e) {
      setJudgeLocks({});
    }
  }, []);

  useEffect(() => {
    loadJudgeLocks();
    const id = setInterval(loadJudgeLocks, 12000);
    return () => clearInterval(id);
  }, [loadJudgeLocks]);

  const setJudgeLock = async (judgeId, locked) => {
    // Writes directly to this judge's own key — no read-merge-write, so 50
    // judges locking around the same time can never overwrite each other.
    setJudgeLocks((prev) => ({ ...prev, [judgeId]: locked }));
    try {
      await storage.set(`jl:${judgeId}`, JSON.stringify(locked), true);
    } catch (e) {}
  };

  const [guidePdf, setGuidePdf] = useState(null);
  const [guideUploading, setGuideUploading] = useState(false);
  const [guideError, setGuideError] = useState("");

  const loadGuidePdf = useCallback(async () => {
    try {
      const res = await storage.get("guide_pdf", true);
      setGuidePdf(res && res.value ? JSON.parse(res.value) : null);
    } catch (e) {
      setGuidePdf(null);
    }
  }, []);

  useEffect(() => {
    loadGuidePdf();
    const id = setInterval(loadGuidePdf, 20000);
    return () => clearInterval(id);
  }, [loadGuidePdf]);

  const uploadGuidePdf = async (file) => {
    if (!file) return;
    if (file.type !== "application/pdf") {
      setGuideError("Lütfen bir PDF dosyası seç.");
      return;
    }
    setGuideUploading(true);
    setGuideError("");
    try {
      const dataUrl = await new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(new Error("Dosya okunamadı"));
        reader.readAsDataURL(file);
      });
      const payload = { name: file.name, data: dataUrl, ts: Date.now() };
      const result = await storage.set("guide_pdf", JSON.stringify(payload), true);
      if (!result) {
        setGuideError("Yükleme başarısız oldu — dosya çok büyük olabilir, daha küçük bir PDF dene.");
      } else {
        setGuidePdf(payload);
      }
    } catch (e) {
      setGuideError("Yükleme sırasında hata oluştu: " + (e?.message || "bilinmeyen hata"));
    }
    setGuideUploading(false);
  };

  const removeGuidePdf = async () => {
    try {
      await storage.delete("guide_pdf", true).catch(() => {});
      setGuidePdf(null);
    } catch (e) {}
  };

  const currentJudgeId =
    session?.role === "judge" ? session.judgeId : session?.role === "admin" ? adminActingJudgeId : null;
  const myLocked = currentJudgeId ? !!judgeLocks[currentJudgeId] : false;
  const scoringLocked = config.locked || myLocked;

  useEffect(() => {
    if (!currentJudgeId || selectedPerson === null) {
      setDraftScores({});
      setDraftNote("");
      return;
    }
    const key = `r:${currentJudgeId}:${selectedPerson}`;
    const existing = ratingsRef.current[key];
    if (existing && existing.scores) {
      setDraftScores({ ...existing.scores });
      setDraftNote(existing.note || "");
    } else {
      const initial = {};
      configRef.current.criteria.forEach((c) => (initial[c.name] = 3));
      setDraftScores(initial);
      setDraftNote("");
    }
    // eslint-disable-next-line
  }, [currentJudgeId, selectedPerson]);

  useEffect(() => {
    if (!currentJudgeId || selectedPresGroup === null) {
      setDraftPresScores({});
      setDraftPresNote("");
      return;
    }
    const key = `pr:${currentJudgeId}:${selectedPresGroup}`;
    const existing = presRatings[key];
    if (existing && existing.scores) {
      setDraftPresScores({ ...existing.scores });
      setDraftPresNote(existing.note || "");
    } else {
      const initial = {};
      configRef.current.presentationCriteria.forEach((c) => (initial[c.name] = 3));
      setDraftPresScores(initial);
      setDraftPresNote("");
    }
  }, [currentJudgeId, selectedPresGroup, presRatings]);

  const savePresRating = async () => {
    if (!currentJudgeId || selectedPresGroup === null || scoringLocked) return;
    setPresSaving(true);
    setPresSaveError("");
    const key = `pr:${currentJudgeId}:${selectedPresGroup}`;
    const judgeObj = config.judges.find((j) => j.id === currentJudgeId);
    const group = config.presentationGroups.find((g) => g.id === selectedPresGroup);
    const payload = {
      scores: draftPresScores,
      note: draftPresNote,
      ts: Date.now(),
      judgeName: judgeObj?.name,
      groupName: group?.name,
    };
    try {
      const result = await setWithRetry(key, JSON.stringify(payload));
      if (!result) {
        setPresSaveError("Kayıt başarısız oldu, lütfen tekrar deneyin.");
      } else {
        setPresRatings((prev) => ({ ...prev, [key]: payload }));
        setPresSavedFlash(true);
        setTimeout(() => setPresSavedFlash(false), 1400);
      }
    } catch (e) {
      setPresSaveError("Kayıt sırasında bir hata oluştu: " + (e?.message || "bilinmeyen hata"));
    }
    setPresSaving(false);
  };

  const handleJudgeLogin = () => {
    setLoginError("");
    const uname = loginUsername.trim().toLowerCase();
    const judge = config.judges.find((j) => j.username.toLowerCase() === uname && j.password === loginPassword);
    if (judge) {
      setSession({ role: "judge", judgeId: judge.id });
      setLoginUsername("");
      setLoginPassword("");
    } else {
      setLoginError("Kullanıcı adı veya şifre hatalı.");
    }
  };

  const handleAdminLogin = () => {
    setLoginError("");
    if (loginAdminPassword === config.adminPassword) {
      setSession({ role: "admin" });
      setLoginAdminPassword("");
      setTab("dashboard");
    } else {
      setLoginError("Yönetici şifresi hatalı.");
    }
  };

  const handleLogout = () => {
    setSession(null);
    setTab("score");
    setSelectedPerson(null);
    setAdminActingJudgeId(null);
    setLoginError("");
  };

  const saveRating = async () => {
    if (!currentJudgeId || selectedPerson === null || scoringLocked) return;
    setSaving(true);
    setSaveError("");
    const key = `r:${currentJudgeId}:${selectedPerson}`;
    const judgeObj = config.judges.find((j) => j.id === currentJudgeId);
    const payload = {
      scores: draftScores,
      note: draftNote,
      ts: Date.now(),
      judgeName: judgeObj?.name,
      personName: config.people[selectedPerson]?.name,
    };
    try {
      const result = await setWithRetry(key, JSON.stringify(payload));
      if (!result) {
        setSaveError("Kayıt başarısız oldu, lütfen tekrar deneyin.");
      } else {
        setRatings((prev) => ({ ...prev, [key]: payload }));
        setSavedFlash(true);
        setTimeout(() => setSavedFlash(false), 1400);
      }
    } catch (e) {
      setSaveError("Kayıt sırasında bir hata oluştu: " + (e?.message || "bilinmeyen hata"));
    }
    setSaving(false);
  };

  const saveConfig = async (next) => {
    setConfig(next);
    setDraftConfig(next);
    try {
      await storage.set("config", JSON.stringify(next), true);
    } catch (e) {}
  };

  const clearAllRatings = async () => {
    if (!window.confirm("Tüm puanlar silinsin mi? Bu işlem geri alınamaz.")) return;
    try {
      const list = await storage.list("r:", true);
      if (list && list.keys) {
        await Promise.all(list.keys.map((k) => storage.delete(k, true).catch(() => {})));
      }
      setRatings({});
    } catch (e) {}
  };

  function personStats(personIdx) {
    const judgeAverages = [];
    config.judges.forEach((j) => {
      const r = ratings[`r:${j.id}:${personIdx}`];
      if (r && r.scores) {
        let sumW = 0;
        let sumSW = 0;
        config.criteria.forEach((c) => {
          const v = r.scores[c.name];
          if (typeof v === "number") {
            const w = c.weight || 1;
            sumSW += v * w;
            sumW += w;
          }
        });
        if (sumW > 0) judgeAverages.push(sumSW / sumW);
      }
    });
    const overall = judgeAverages.length
      ? judgeAverages.reduce((a, b) => a + b, 0) / judgeAverages.length
      : null;
    return { overall, count: judgeAverages.length, total: config.judges.length };
  }

  function notesForPerson(personIdx) {
    const notes = [];
    config.judges.forEach((j) => {
      const r = ratings[`r:${j.id}:${personIdx}`];
      if (r && r.note && r.note.trim()) {
        notes.push({ judgeName: j.name, note: r.note, ts: r.ts });
      }
    });
    return notes.sort((a, b) => b.ts - a.ts);
  }

  function presGroupStats(groupId) {
    const judgeAverages = [];
    config.judges.forEach((j) => {
      const r = presRatings[`pr:${j.id}:${groupId}`];
      if (r && r.scores) {
        let sumW = 0;
        let sumSW = 0;
        config.presentationCriteria.forEach((c) => {
          const v = r.scores[c.name];
          if (typeof v === "number") {
            const w = c.weight || 1;
            sumSW += v * w;
            sumW += w;
          }
        });
        if (sumW > 0) judgeAverages.push(sumSW / sumW);
      }
    });
    const overall = judgeAverages.length
      ? judgeAverages.reduce((a, b) => a + b, 0) / judgeAverages.length
      : null;
    return { overall, count: judgeAverages.length, total: config.judges.length };
  }

  function notesForPresGroup(groupId) {
    const notes = [];
    config.judges.forEach((j) => {
      const r = presRatings[`pr:${j.id}:${groupId}`];
      if (r && r.note && r.note.trim()) {
        notes.push({ judgeName: j.name, note: r.note, ts: r.ts });
      }
    });
    return notes.sort((a, b) => b.ts - a.ts);
  }

  const presGroupsFull = React.useMemo(
    () =>
      config.presentationGroups.map((g) => ({
        idx: g.id,
        name: g.name,
        photo: null,
        notes: notesForPresGroup(g.id),
        ...presGroupStats(g.id),
      })),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [config.presentationGroups, config.judges, config.presentationCriteria, presRatings]
  );

  const peopleFull = React.useMemo(
    () =>
      config.people.map((p, idx) => ({
        idx,
        name: p.name,
        groupId: p.groupId,
        photo: p.photo,
        notes: notesForPerson(idx),
        ...personStats(idx),
      })),
    // eslint-disable-next-line react-hooks/exhaustive-deps
    [config.people, config.judges, config.criteria, ratings]
  );

  const sortByScore = (arr) =>
    [...arr].sort((a, b) => {
      if (a.overall === null && b.overall === null) return 0;
      if (a.overall === null) return 1;
      if (b.overall === null) return -1;
      return b.overall - a.overall;
    });

  const knownGroupIds = new Set(config.groups.map((g) => g.id));
  const hasUnassigned = peopleFull.some((p) => !knownGroupIds.has(p.groupId));
  const groupOptions = [...config.groups, ...(hasUnassigned ? [{ id: "__unassigned", name: "Grupsuz" }] : [])];

  const groupBlocks = groupOptions.map((g) => ({
    ...g,
    members: sortByScore(
      peopleFull.filter((p) => (g.id === "__unassigned" ? !knownGroupIds.has(p.groupId) : p.groupId === g.id))
    ),
  }));

  const scoreVisiblePeopleBase =
    scoreGroupFilter === "all"
      ? config.people.map((p, i) => ({ ...p, idx: i }))
      : config.people
          .map((p, i) => ({ ...p, idx: i }))
          .filter((p) =>
            scoreGroupFilter === "__unassigned" ? !knownGroupIds.has(p.groupId) : p.groupId === scoreGroupFilter
          );

  const scoreVisiblePeople = showOnlyUnscored
    ? scoreVisiblePeopleBase.filter((p) => !ratings[`r:${currentJudgeId}:${p.idx}`])
    : scoreVisiblePeopleBase;

  const myDoneCount = currentJudgeId
    ? config.people.filter((_, idx) => !!ratings[`r:${currentJudgeId}:${idx}`]).length
    : 0;
  const myTotalCount = config.people.length;
  const myPct = myTotalCount > 0 ? Math.round((myDoneCount / myTotalCount) * 100) : 0;

  const recentActivity = Object.entries(ratings)
    .filter(([, r]) => r && r.ts)
    .map(([k, r]) => {
      const parts = k.split(":");
      const jId = parts[1];
      const pIdx = Number(parts[2]);
      return {
        ts: r.ts,
        judgeName: config.judges.find((j) => j.id === jId)?.name ?? r.judgeName,
        personName: config.people[pIdx]?.name ?? r.personName,
      };
    })
    .sort((a, b) => b.ts - a.ts)
    .slice(0, 8);

  const totalPossibleRatings = config.people.length * config.judges.length;
  const totalDoneRatings = Object.values(ratings).filter((r) => r && r.scores).length;
  const completionPct = totalPossibleRatings > 0 ? Math.round((totalDoneRatings / totalPossibleRatings) * 100) : 0;

  const currentJudgeObj = config.judges.find((j) => j.id === currentJudgeId);
  const shareUrl = (config.appUrl || "").trim();

  function personCriteriaAverages(personIdx) {
    const sums = {};
    const counts = {};
    config.judges.forEach((j) => {
      const r = ratings[`r:${j.id}:${personIdx}`];
      if (r && r.scores) {
        Object.entries(r.scores).forEach(([c, v]) => {
          if (typeof v === "number") {
            sums[c] = (sums[c] || 0) + v;
            counts[c] = (counts[c] || 0) + 1;
          }
        });
      }
    });
    const avgs = {};
    config.criteria.forEach((c) => {
      avgs[c.name] = counts[c.name] ? Number((sums[c.name] / counts[c.name]).toFixed(2)) : "";
    });
    return avgs;
  }

  const exportToExcel = () => {
    setExportError("");
    try {
      const groupName = (gid) => config.groups.find((g) => g.id === gid)?.name ?? "Grupsuz";
      const sortedPeople = sortByScore(peopleFull);

      const summaryRows = sortedPeople.map((p, i) => {
        const critAvgs = personCriteriaAverages(p.idx);
        const row = { Sıra: i + 1, Yarışmacı: p.name, Grup: groupName(p.groupId) };
        config.criteria.forEach((c) => (row[c.weight !== 1 ? `${c.name} (x${c.weight})` : c.name] = critAvgs[c.name]));
        row["Ortalama"] = p.overall !== null ? Number(p.overall.toFixed(2)) : "";
        row["Hakem Sayısı"] = `${p.count}/${p.total}`;
        return row;
      });
      if (summaryRows.length === 0) {
        summaryRows.push({ Sıra: "", Yarışmacı: "(yarışmacı yok)", Grup: "" });
      }

      const detailRows = [];
      Object.entries(ratings).forEach(([k, r]) => {
        if (!r || !r.scores) return;
        const parts = k.split(":");
        const jId = parts[1];
        const pIdx = Number(parts[2]);
        const judge = config.judges.find((j) => j.id === jId);
        const person = config.people[pIdx];
        const row = {
          Hakem: judge?.name ?? r.judgeName ?? "",
          Yarışmacı: person?.name ?? r.personName ?? "",
          Grup: person ? groupName(person.groupId) : "",
        };
        config.criteria.forEach((c) => (row[c.weight !== 1 ? `${c.name} (x${c.weight})` : c.name] = typeof r.scores[c.name] === "number" ? r.scores[c.name] : ""));
        let sumW = 0;
        let sumSW = 0;
        config.criteria.forEach((c) => {
          const v = r.scores[c.name];
          if (typeof v === "number") {
            sumSW += v * (c.weight || 1);
            sumW += c.weight || 1;
          }
        });
        row["Ortalama"] = sumW > 0 ? Number((sumSW / sumW).toFixed(2)) : "";
        row["Tarih"] = r.ts ? new Date(r.ts).toLocaleString("tr-TR") : "";
        row["Not"] = r.note || "";
        detailRows.push(row);
      });
      detailRows.sort((a, b) => (a.Tarih < b.Tarih ? 1 : -1));
      if (detailRows.length === 0) {
        detailRows.push({ Hakem: "(henüz puan girilmedi)" });
      }

      const wb = XLSX.utils.book_new();
      const ws1 = XLSX.utils.json_to_sheet(summaryRows);
      const ws2 = XLSX.utils.json_to_sheet(detailRows);
      XLSX.utils.book_append_sheet(wb, ws1, "Sonuçlar");
      XLSX.utils.book_append_sheet(wb, ws2, "Detay");

      const wbout = XLSX.write(wb, { bookType: "xlsx", type: "array" });
      const blob = new Blob([wbout], { type: "application/octet-stream" });
      const url = URL.createObjectURL(blob);
      const stamp = new Date().toISOString().slice(0, 10);
      const a = document.createElement("a");
      a.href = url;
      a.download = `puanlama-sonuclari-${stamp}.xlsx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    } catch (e) {
      setExportError("Excel oluşturulamadı: " + (e?.message || "bilinmeyen hata"));
    }
  };

  if (showReport) {
    return (
      <ReportView
        config={config}
        rankedPeople={sortByScore(peopleFull)}
        personCriteriaAverages={personCriteriaAverages}
        completionPct={completionPct}
        totalDoneRatings={totalDoneRatings}
        totalPossibleRatings={totalPossibleRatings}
        onClose={() => setShowReport(false)}
      />
    );
  }

  return (
    <div className="sb-root">
      <div className="sb-mesh" aria-hidden="true" />
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600;700&display=swap');

        html, body {
          background: #F2F2F0;
          margin: 0;
          min-height: 100%;
        }
        .sb-mesh {
          position: fixed; inset: 0; z-index: 0; pointer-events: none; overflow: hidden;
        }
        .sb-mesh::before, .sb-mesh::after {
          content: ''; position: absolute; border-radius: 50%; filter: blur(100px); opacity: 0.16; will-change: transform;
        }
        .sb-mesh::before {
          width: 660px; height: 660px;
          background: radial-gradient(circle at 30% 30%, #e63946, transparent 60%);
          top: -180px; left: -120px;
          animation: sb-drift1 26s cubic-bezier(0.32,0.72,0,1) infinite alternate;
        }
        .sb-mesh::after {
          width: 500px; height: 500px;
          background: radial-gradient(circle at 60% 60%, #eab308, transparent 60%);
          top: 140px; right: -100px;
          animation: sb-drift2 34s cubic-bezier(0.32,0.72,0,1) infinite alternate;
        }
        @keyframes sb-drift1 { from { transform: translate3d(0,0,0); } to { transform: translate3d(70px, 50px, 0); } }
        @keyframes sb-drift2 { from { transform: translate3d(0,0,0); } to { transform: translate3d(-50px, 35px, 0); } }
        @media (prefers-reduced-motion: reduce) { .sb-mesh::before, .sb-mesh::after { animation: none; } }

        .sb-root {
          --ink: #F2F2F0;
          --panel: rgba(255,255,255,0.65);
          --panel-2: rgba(255,255,255,0.42);
          --line: rgba(10,10,10,0.07);
          --line-strong: rgba(10,10,10,0.11);
          --innerlight: rgba(255,255,255,0.65);
          --amber: #eab308;
          --amber-dim: #a16207;
          --teal: #e63946;
          --text: #0A0A0A;
          --text-dim: #57575A;
          --danger: #dc2626;
          --ease: cubic-bezier(0.32, 0.72, 0, 1);
          --ease-spring: cubic-bezier(0.16, 1.16, 0.3, 1);
          --shell-radius: 1.25rem;
          --core-radius: calc(1.25rem - 0.3rem);
          font-family: 'Plus Jakarta Sans', system-ui, sans-serif;
          -webkit-font-smoothing: antialiased;
          background: transparent;
          color: var(--text);
          min-height: 100vh;
          padding: 28px 18px 60px;
          box-sizing: border-box;
          position: relative;
          z-index: 1;
        }
        .sb-root * { box-sizing: border-box; }
        .sb-wrap { max-width: 880px; margin: 0 auto; position: relative; z-index: 1; }

        .sb-welcome {
          background: var(--panel);
          border: 1px solid var(--line);
          border-radius: 12px;
          padding: 14px 0;
          margin-bottom: 18px;
          overflow: hidden;
        }
        .sb-welcome-label {
          font-family: 'Oswald', sans-serif;
          font-weight: 600;
          font-size: 14px;
          text-transform: uppercase;
          letter-spacing: 0.06em;
          color: var(--amber-dim);
          padding: 0 16px 10px;
        }
        .sb-marquee { overflow: hidden; white-space: nowrap; -webkit-mask-image: linear-gradient(90deg, transparent, #000 6%, #000 94%, transparent); mask-image: linear-gradient(90deg, transparent, #000 6%, #000 94%, transparent); }
        .sb-marquee-track {
          display: inline-block;
          white-space: nowrap;
          animation: sb-marquee-scroll 22s linear infinite;
        }
        .sb-marquee-track span {
          display: inline-block;
          font-family: 'Oswald', sans-serif;
          font-weight: 500;
          font-size: 15px;
          color: var(--text);
          padding: 0 20px;
          position: relative;
        }
        .sb-marquee-track span::after {
          content: "•";
          color: var(--teal);
          margin-left: 20px;
        }
        @keyframes sb-marquee-scroll {
          from { transform: translateX(0); }
          to { transform: translateX(-50%); }
        }
        @media (prefers-reduced-motion: reduce) {
          .sb-marquee-track { animation: none; }
        }

        .sb-header { margin-bottom: 22px; display: flex; flex-direction: column; align-items: center; gap: 14px; }
        .sb-header-center { display: flex; flex-direction: column; align-items: center; text-align: center; }

        .sb-hero-shell {
          width: 100%;
          padding: 8px;
          border-radius: calc(var(--shell-radius) + 4px);
          background: rgba(255,255,255,0.35);
          box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.5),
            0 0 0 1px var(--line),
            0 20px 50px -20px rgba(10,10,10,0.15);
        }
        .sb-hero-core {
          position: relative;
          overflow: hidden;
          display: flex;
          flex-direction: column;
          align-items: center;
          text-align: center;
          padding: 38px 28px 28px;
          border-radius: var(--shell-radius);
          background: rgba(255,255,255,0.7);
          backdrop-filter: saturate(180%) blur(30px);
          -webkit-backdrop-filter: saturate(180%) blur(30px);
          box-shadow: inset 0 1px 0 rgba(255,255,255,0.8);
        }
        .sb-hero-accent {
          position: absolute;
          top: -60px; right: -60px;
          width: 180px; height: 180px;
          border-radius: 50%;
          background: radial-gradient(circle, rgba(234,179,8,0.18), transparent 70%);
          pointer-events: none;
        }
        .sb-hero-badge {
          display: inline-flex;
          padding: 5px 14px;
          border-radius: 999px;
          background: var(--text);
          color: #F2F2F0;
          font-family: 'JetBrains Mono', monospace;
          font-size: 10px;
          font-weight: 700;
          letter-spacing: 0.18em;
          margin-bottom: 14px;
          box-shadow: inset 0 1px 0 rgba(255,255,255,0.1);
        }
        .sb-hero-line {
          display: flex;
          gap: 6px;
          margin-top: 16px;
        }
        .sb-hero-dot {
          width: 5px; height: 5px;
          border-radius: 50%;
          background: var(--line-strong);
        }
        .sb-hero-dot:nth-child(2) { background: var(--amber); }
        .sb-eyebrow {
          font-family: 'JetBrains Mono', monospace;
          font-size: 10px;
          letter-spacing: 0.22em;
          text-transform: uppercase;
          color: var(--text-dim);
          display: inline-flex;
          align-items: center;
          gap: 8px;
          margin-bottom: 10px;
          padding: 6px 14px;
          border-radius: 999px;
          background: rgba(255,255,255,0.6);
          box-shadow: inset 0 0 0 1px var(--line);
        }
        }
        .sb-eyebrow .dot {
          width: 7px; height: 7px; border-radius: 50%;
          background: var(--amber);
          box-shadow: 0 0 0 0 rgba(234,179,8,0.6);
          animation: pulse 1.8s infinite;
        }
        @keyframes pulse {
          0% { box-shadow: 0 0 0 0 rgba(234,179,8,0.55); }
          70% { box-shadow: 0 0 0 8px rgba(234,179,8,0); }
          100% { box-shadow: 0 0 0 0 rgba(234,179,8,0); }
        }
        .sb-title {
          font-family: 'Plus Jakarta Sans', system-ui, sans-serif;
          font-weight: 800;
          font-size: 32px;
          letter-spacing: -0.03em;
          margin: 0 0 6px;
          line-height: 1.1;
        }
        .sb-sub { color: var(--text-dim); font-size: 14px; font-weight: 500; letter-spacing: 0.02em; margin-top: 2px; }

        .sb-sessionbox { display: flex; align-items: center; gap: 10px; }
        .sb-sessionname { font-size: 12px; color: var(--text-dim); font-family: 'JetBrains Mono', monospace; text-align: center; }
        .sb-sessionname b { color: var(--text); }
        .sb-logoutbtn {
          display: flex; align-items: center; gap: 6px;
          padding: 7px 14px; border-radius: 999px;
          border: 1px solid var(--line); background: rgba(255,255,255,0.5);
          color: var(--text-dim); font-size: 12px; font-weight: 600; cursor: pointer;
          transition: background 300ms var(--ease), color 200ms var(--ease);
        }
        .sb-logoutbtn:hover { background: rgba(220,38,38,0.08); color: var(--danger); border-color: rgba(220,38,38,0.2); }

        .sb-tabs {
          display: flex;
          gap: 6px;
          padding: 6px;
          border-radius: 999px;
          background: rgba(255,255,255,0.42);
          backdrop-filter: saturate(160%) blur(22px);
          -webkit-backdrop-filter: saturate(160%) blur(22px);
          box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.5),
            0 0 0 1px var(--line),
            0 8px 24px -12px rgba(10,10,10,0.12);
          margin: 20px 0 22px;
        }
        .sb-tab {
          flex: 1;
          display: flex;
          align-items: center;
          justify-content: center;
          gap: 7px;
          padding: 10px 8px;
          border-radius: 999px;
          border: none;
          background: transparent;
          color: var(--text-dim);
          font-family: 'Plus Jakarta Sans', system-ui, sans-serif;
          font-weight: 600;
          font-size: 13px;
          cursor: pointer;
          transition: background 300ms var(--ease), color 200ms var(--ease), transform 300ms var(--ease-spring);
        }
        .sb-tab:hover { background: rgba(255,255,255,0.5); }
        .sb-tab:active { transform: scale(0.97); }
        .sb-tab.active {
          background: rgba(255,255,255,0.75);
          color: var(--text);
          box-shadow: inset 0 1px 0 rgba(255,255,255,0.7), 0 1px 3px rgba(10,10,10,0.08);
        }
        .sb-tab:focus-visible { outline: 2px solid var(--teal); outline-offset: 2px; }

        .sb-card {
          background: var(--panel);
          backdrop-filter: saturate(160%) blur(22px);
          -webkit-backdrop-filter: saturate(160%) blur(22px);
          border: 1px solid var(--line);
          border-radius: var(--shell-radius);
          padding: 18px;
          margin-bottom: 14px;
          box-shadow:
            inset 0 1px 0 var(--innerlight),
            0 0 0 1px var(--line),
            0 12px 36px -18px rgba(10,10,10,0.12),
            0 4px 12px -6px rgba(10,10,10,0.06);
          transition: box-shadow 300ms var(--ease), transform 300ms var(--ease);
        }
        .sb-card:hover {
          box-shadow:
            inset 0 1px 0 var(--innerlight),
            0 0 0 1px var(--line),
            0 16px 48px -18px rgba(10,10,10,0.16),
            0 6px 18px -8px rgba(10,10,10,0.08);
        }
        .sb-label {
          font-family: 'JetBrains Mono', monospace;
          font-size: 10.5px;
          letter-spacing: 0.1em;
          color: var(--text-dim);
          text-transform: uppercase;
          margin-bottom: 10px;
        }
        .sb-section-divider {
          display: flex;
          align-items: center;
          gap: 8px;
          margin: 26px 0 14px;
          padding-top: 18px;
          border-top: 1px solid var(--line);
          font-family: 'JetBrains Mono', monospace;
          font-size: 12px;
          letter-spacing: 0.08em;
          text-transform: uppercase;
          color: var(--text-dim);
        }
        .sb-groupheader {
          display: flex;
          align-items: baseline;
          justify-content: space-between;
          margin-bottom: 12px;
        }
        .sb-groupheader .gname {
          font-family: 'Oswald', sans-serif;
          font-weight: 600;
          font-size: 15px;
          text-transform: uppercase;
          letter-spacing: 0.03em;
          color: var(--teal);
        }
        .sb-groupheader .gcount {
          font-family: 'JetBrains Mono', monospace;
          font-size: 11px;
          color: var(--text-dim);
        }

        .sb-chiprow { display: flex; flex-wrap: wrap; gap: 8px; }
        .sb-chip {
          padding: 9px 16px;
          border-radius: 999px;
          border: 1px solid var(--line);
          background: rgba(255,255,255,0.5);
          color: var(--text);
          font-size: 13.5px;
          font-weight: 500;
          cursor: pointer;
          transition: background 300ms var(--ease), border-color 200ms var(--ease), transform 300ms var(--ease-spring);
        }
        .sb-chip:hover { background: rgba(255,255,255,0.8); }
        .sb-chip:active { transform: scale(0.97); }
        .sb-chip.active { border-color: var(--text); color: var(--text); background: rgba(255,255,255,0.85); box-shadow: inset 0 1px 0 rgba(255,255,255,0.6); }
        .sb-chip:focus-visible { outline: 2px solid var(--teal); outline-offset: 2px; }

        .sb-peoplegrid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(180px, 1fr));
          gap: 12px;
        }
        .sb-personcard {
          position: relative;
          aspect-ratio: 1 / 1;
          padding: 16px 14px;
          border-radius: var(--shell-radius);
          border: 1px solid var(--line);
          background: rgba(255,255,255,0.5);
          backdrop-filter: blur(12px);
          -webkit-backdrop-filter: blur(12px);
          cursor: pointer;
          text-align: center;
          display: flex;
          flex-direction: column;
          align-items: center;
          justify-content: center;
          gap: 4px;
          box-shadow: inset 0 1px 0 var(--innerlight), 0 2px 8px rgba(10,10,10,0.06);
          transition: transform 300ms var(--ease-spring), box-shadow 300ms var(--ease), border-color 200ms var(--ease);
        }
        .sb-personcard:hover { transform: translateY(-2px); box-shadow: inset 0 1px 0 var(--innerlight), 0 8px 24px rgba(10,10,10,0.1); }
        .sb-personcard:active { transform: scale(0.98); }
        .sb-personcard.active { border-color: var(--teal); box-shadow: inset 0 1px 0 var(--innerlight), 0 0 0 1px var(--teal), 0 8px 24px rgba(230,57,70,0.12); }
        .sb-personcard .pname { font-weight: 600; font-size: 15px; }
        .sb-personcard .pstatus {
          display: flex; align-items: center; justify-content: center; gap: 5px;
          margin-top: 6px; font-size: 11.5px; color: var(--text-dim);
          font-family: 'JetBrains Mono', monospace;
        }
        .sb-personcard .check { color: var(--teal); }
        .sb-personcard .ptop { display: flex; flex-direction: column; align-items: center; gap: 10px; margin-bottom: 2px; }
        .sb-personcard .pmeta { font-size: 11px; color: var(--text-dim); margin-top: 2px; line-height: 1.4; text-align: center; }

        .sb-avatar {
          border-radius: 50%;
          object-fit: cover;
          flex-shrink: 0;
        }
        .sb-avatar-fallback {
          display: flex; align-items: center; justify-content: center;
          background: var(--line);
          color: var(--text-dim);
          font-weight: 700;
          font-family: 'Oswald', sans-serif;
        }

        .sb-profilehead {
          display: flex; align-items: center; gap: 14px;
          margin-bottom: 16px;
          padding-bottom: 14px;
          border-bottom: 1px solid var(--line);
        }
        .sb-profilehead .pname { font-family: 'Oswald', sans-serif; font-weight: 600; font-size: 17px; }
        .sb-profilehead .pmeta { font-size: 12.5px; color: var(--text-dim); margin-top: 3px; line-height: 1.5; }

        .sb-photoupload { display: flex; align-items: center; gap: 10px; }
        .sb-photoupload label {
          display: flex; align-items: center; justify-content: center;
          width: 84px; height: 84px; border-radius: 16px;
          border: 1px dashed var(--line);
          color: var(--text-dim);
          cursor: pointer;
          flex-shrink: 0;
          overflow: hidden;
        }
        .sb-photoupload label:hover { border-color: var(--amber-dim); color: var(--amber-dim); }
        .sb-photoupload input[type="file"] { display: none; }

        .sb-slider-row { margin-bottom: 22px; }
        .sb-slider-head { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
        .sb-crit-name { font-weight: 600; font-size: 14px; }
        .sb-critdesc { font-size: 12px; color: var(--text-dim); line-height: 1.5; margin: -6px 0 10px; }
        .sb-weightbadge {
          display: inline-block;
          margin-left: 7px;
          padding: 1px 7px;
          border-radius: 999px;
          background: rgba(230,57,70,0.18);
          color: #e63946;
          font-family: 'JetBrains Mono', monospace;
          font-size: 10.5px;
          font-weight: 700;
          vertical-align: middle;
        }
        .sb-crit-val {
          font-family: 'JetBrains Mono', monospace;
          color: #000000;
          background: var(--amber);
          font-weight: 700;
          font-size: 15px;
          min-width: 34px;
          text-align: center;
          padding: 3px 0;
          border-radius: 8px;
          box-shadow: 0 2px 10px rgba(234,179,8,0.35);
          animation: sb-pop 0.18s ease;
        }
        @keyframes sb-pop {
          0% { transform: scale(1.35); }
          100% { transform: scale(1); }
        }
        .sb-scoreboxes { display: flex; gap: 6px; flex-wrap: wrap; }
        .sb-scorebox {
          flex: 1;
          min-width: 30px;
          padding: 12px 0;
          text-align: center;
          border-radius: 9px;
          border: 1px solid var(--line);
          background: var(--panel-2);
          color: var(--text);
          font-family: 'JetBrains Mono', monospace;
          font-weight: 700;
          font-size: 14px;
          cursor: pointer;
          transition: transform 0.1s ease, background 0.1s ease, border-color 0.1s ease, box-shadow 0.1s ease;
        }
        .sb-scorebox:hover { border-color: var(--teal); }
        .sb-scorebox.active {
          background: var(--amber);
          color: #000000;
          border-color: var(--amber-dim);
          box-shadow: 0 3px 12px rgba(234,179,8,0.4);
          transform: scale(1.08);
        }
        .sb-scorebox:focus-visible { outline: 2px solid var(--text); outline-offset: 2px; }
        .sb-notearea {
          width: 100%;
          background: var(--panel-2);
          border: 1px solid var(--line);
          border-radius: 10px;
          padding: 12px;
          color: var(--text);
          font-family: 'Inter', sans-serif;
          font-size: 13.5px;
          resize: vertical;
          min-height: 64px;
        }
        .sb-notearea:focus-visible { outline: 2px solid var(--amber); }
        .sb-notelist { display: flex; flex-direction: column; gap: 16px; max-height: 420px; overflow-y: auto; }
        .sb-notegroup { display: flex; flex-direction: column; gap: 6px; }
        .sb-notegroup-head {
          display: flex; align-items: center; gap: 8px;
          font-family: 'Oswald', sans-serif; font-weight: 600; font-size: 14px;
          color: var(--text);
          padding-bottom: 6px;
          border-bottom: 1px solid var(--line);
        }
        .sb-noteitem { background: var(--panel-2); border: 1px solid var(--line); border-radius: 10px; padding: 10px 12px; }
        .sb-noteitem-head { font-size: 12px; color: var(--text-dim); margin-bottom: 4px; }
        .sb-noteitem-head b { color: var(--amber-dim); }
        .sb-noteitem-text { font-size: 13.5px; color: var(--text); line-height: 1.5; }

        .sb-savebtn {
          width: 100%;
          padding: 13px;
          border-radius: 999px;
          border: none;
          background: var(--text);
          color: #F2F2F0;
          font-weight: 700;
          font-size: 14px;
          cursor: pointer;
          display: flex; align-items: center; justify-content: center; gap: 8px;
          box-shadow: inset 0 1px 0 rgba(255,255,255,0.12), 0 2px 8px rgba(10,10,10,0.15);
          transition: transform 300ms var(--ease), box-shadow 300ms var(--ease);
        }
        .sb-savebtn:hover { box-shadow: inset 0 1px 0 rgba(255,255,255,0.12), 0 4px 16px rgba(10,10,10,0.2); }
        .sb-savebtn:active { transform: scale(0.98); }
        .sb-savebtn:disabled { opacity: 0.5; cursor: default; transform: none; }
        .sb-savebtn:focus-visible { outline: 2px solid var(--teal); outline-offset: 2px; }
        .sb-savedflash { color: var(--teal); font-size: 12.5px; text-align: center; margin-top: 8px; font-family: 'JetBrains Mono', monospace; }

        .sb-ticker {
          overflow: hidden;
          border: 1px solid var(--line);
          border-radius: 10px;
          background: var(--panel);
          padding: 10px 14px;
          margin-bottom: 16px;
          display: flex;
          gap: 10px;
          align-items: center;
        }
        .sb-ticker-items { display: flex; gap: 22px; overflow-x: auto; font-family: 'JetBrains Mono', monospace; font-size: 11.5px; color: var(--text-dim); }
        .sb-ticker-items b { color: var(--amber-dim); font-weight: 700; }

        .sb-timerbar {
          display: flex; align-items: center; gap: 12px;
          background: var(--panel);
          border: 1px solid var(--line);
          border-radius: 12px;
          padding: 10px 16px;
          margin-bottom: 16px;
          flex-wrap: wrap;
        }
        .sb-timerdigits {
          font-family: 'JetBrains Mono', monospace;
          font-weight: 700;
          font-size: 20px;
          color: var(--text);
          letter-spacing: 0.03em;
        }
        .sb-timercontrols { display: flex; align-items: center; gap: 6px; margin-left: auto; }
        .sb-guidebtn {
          display: flex; align-items: center; gap: 6px;
          padding: 7px 13px; border-radius: 999px;
          border: 1px solid var(--line); background: rgba(255,255,255,0.5);
          color: var(--teal); font-size: 12.5px; font-weight: 600;
          text-decoration: none; cursor: pointer;
          margin-left: auto;
          transition: background 200ms var(--ease), border-color 200ms var(--ease);
        }
        .sb-guidebtn:hover { background: rgba(230,57,70,0.08); border-color: rgba(230,57,70,0.25); }
        .sb-timerunit { font-size: 11px; color: var(--text-dim); }
        .sb-timerbtn {
          display: flex; align-items: center; justify-content: center;
          width: 32px; height: 32px; border-radius: 8px;
          border: 1px solid var(--line); background: var(--panel-2);
          color: var(--amber-dim); cursor: pointer;
        }
        .sb-timerbtn:hover { border-color: var(--amber-dim); }
        .sb-progresscard { padding: 16px 18px; }
        .sb-myprogress { margin-bottom: 14px; }
        .sb-progresshead {
          display: flex; justify-content: space-between; align-items: baseline;
          font-size: 13px; color: var(--text-dim); margin-bottom: 10px;
        }
        .sb-progresspct {
          font-family: 'JetBrains Mono', monospace;
          font-weight: 700;
          font-size: 16px;
          color: var(--amber-dim);
        }
        .sb-progresstrack {
          width: 100%;
          height: 10px;
          border-radius: 999px;
          background: var(--line);
          overflow: hidden;
        }
        .sb-progressfill {
          height: 100%;
          border-radius: 999px;
          background: linear-gradient(90deg, #9c1f29, #e63946, #f2c230);
          transition: width 0.4s ease;
        }
        .sb-progresssub {
          margin-top: 8px;
          font-size: 11.5px;
          color: var(--text-dim);
          font-family: 'JetBrains Mono', monospace;
        }
        .sb-viewtoggle { display: flex; gap: 6px; margin-bottom: 12px; flex-wrap: wrap; }

        .sb-viewbtn {
          display: flex; align-items: center; gap: 6px;
          padding: 8px 13px; border-radius: 8px;
          border: 1px solid var(--line); background: var(--panel);
          color: var(--text-dim); font-size: 12.5px; font-weight: 600; cursor: pointer;
        }
        .sb-viewbtn.active { border-color: var(--amber-dim); color: var(--amber-dim); background: rgba(234,179,8,0.08); }
        .sb-viewbtn:focus-visible { outline: 2px solid var(--amber); outline-offset: 2px; }
        .sb-empty { text-align: center; color: var(--text-dim); font-size: 13px; padding: 30px 0; }

        .sb-rankrow {
          display: flex;
          align-items: flex-start;
          gap: 14px;
          padding: 12px 6px;
          border-bottom: 1px solid var(--line);
        }
        .sb-rankrow:last-child { border-bottom: none; }
        .sb-rankpos {
          font-family: 'Oswald', sans-serif;
          font-size: 16px;
          font-weight: 600;
          color: var(--text-dim);
          width: 26px;
          text-align: center;
          margin-top: 6px;
        }
        .sb-rankrow:nth-child(1) .sb-rankpos { color: var(--amber-dim); }
        .sb-rankname { flex: 1; font-weight: 600; font-size: 14.5px; }
        .sb-rankmeta { font-size: 11px; color: var(--text-dim); font-family: 'JetBrains Mono', monospace; margin-top: 2px; }
        .sb-rownotes { display: flex; flex-direction: column; gap: 4px; margin-top: 7px; }
        .sb-rownote {
          font-size: 12px;
          color: var(--text-dim);
          line-height: 1.5;
          background: var(--panel-2);
          border-radius: 7px;
          padding: 6px 9px;
        }
        .sb-rownote b { color: var(--amber-dim); font-weight: 600; }
        .sb-rankrow .flip-group { margin-top: 6px; }

        .flip-group { display: flex; gap: 2px; }
        .flip-digit, .flip-dot {
          font-family: 'JetBrains Mono', monospace;
          font-weight: 700;
          background: #000000;
          border: 1px solid var(--line);
          border-radius: 4px;
          color: #eab308;
          text-align: center;
        }
        .flip-lg .flip-digit { min-width: 22px; font-size: 20px; padding: 3px 2px; }
        .flip-lg .flip-dot { min-width: 8px; font-size: 20px; padding: 3px 0; border: none; background: transparent; }
        .flip-sm .flip-digit { min-width: 16px; font-size: 13px; padding: 2px 1px; }
        .flip-sm .flip-dot { min-width: 5px; font-size: 13px; padding: 2px 0; border: none; background: transparent; }

        .sb-settings-list { display: flex; flex-direction: column; gap: 8px; }
        .sb-settings-item { display: flex; gap: 8px; align-items: center; flex-wrap: wrap; }
        .sb-input {
          flex: 1;
          min-width: 90px;
          background: rgba(255,255,255,0.55);
          border: 1px solid var(--line);
          border-radius: 12px;
          padding: 11px 14px;
          color: var(--text);
          font-size: 14px;
          font-family: 'Plus Jakarta Sans', system-ui, sans-serif;
          box-shadow: inset 0 2px 4px rgba(10,10,10,0.04);
          transition: border-color 200ms var(--ease), box-shadow 200ms var(--ease);
        }
        .sb-input::placeholder { color: var(--text-dim); opacity: 0.6; }
        .sb-input:focus { outline: none; border-color: var(--teal); box-shadow: inset 0 2px 4px rgba(10,10,10,0.04), 0 0 0 3px rgba(230,57,70,0.1); }
        .sb-select {
          background: var(--panel-2);
          border: 1px solid var(--line);
          border-radius: 8px;
          padding: 9px 8px;
          color: var(--text);
          font-size: 12.5px;
          font-family: 'Inter', sans-serif;
          max-width: 118px;
        }
        .sb-select:focus-visible { outline: 2px solid var(--amber); }
        .sb-iconbtn {
          border: 1px solid var(--line);
          background: var(--panel-2);
          color: var(--text-dim);
          border-radius: 8px;
          width: 34px; height: 34px;
          flex-shrink: 0;
          display: flex; align-items: center; justify-content: center;
          cursor: pointer;
        }
        .sb-iconbtn:hover { color: var(--danger); border-color: var(--danger); }
        .sb-iconbtn-active { color: var(--amber-dim); border-color: var(--amber-dim); background: rgba(234,179,8,0.12); }
        .sb-iconbtn-active:hover { color: var(--amber-dim); border-color: var(--amber-dim); }

        .sb-lockrow {
          display: flex;
          align-items: center;
          justify-content: space-between;
          gap: 14px;
          flex-wrap: wrap;
        }
        .sb-lockstatus {
          display: flex;
          align-items: center;
          gap: 7px;
          font-weight: 600;
          font-size: 14px;
          color: var(--text);
        }
        .sb-lockbanner {
          display: flex;
          align-items: center;
          gap: 8px;
          background: rgba(220,38,38,0.1);
          border: 1px solid rgba(220,38,38,0.35);
          color: var(--danger);
          border-radius: 9px;
          padding: 10px 13px;
          font-size: 12.5px;
          margin-bottom: 16px;
        }
        .sb-fireworks-canvas {
          position: fixed;
          inset: 0;
          width: 100%;
          height: 100%;
          pointer-events: none;
          z-index: 50;
        }
        .sb-katilanlar { padding-top: 6px; }
        .sb-katilanlar-group { margin-bottom: 34px; }
        .sb-katilanlar-group:last-child { margin-bottom: 0; }
        .sb-katilanlar-title {
          font-family: 'Oswald', sans-serif;
          font-weight: 700;
          font-size: 28px;
          text-transform: uppercase;
          letter-spacing: 0.03em;
          text-align: center;
          color: var(--text);
          margin-bottom: 26px;
        }

        .sb-showcase-title {
          font-family: 'Oswald', sans-serif;
          font-weight: 700;
          font-size: 20px;
          text-transform: uppercase;
          letter-spacing: 0.03em;
          color: var(--text);
          margin-bottom: 18px;
          text-align: center;
        }
        .sb-showcasegrid {
          display: grid;
          grid-template-columns: repeat(auto-fill, minmax(150px, 1fr));
          gap: 18px;
        }
        .sb-showcasecard {
          display: flex;
          flex-direction: column;
          align-items: center;
          text-align: center;
          gap: 10px;
        }
        .sb-showcasename { font-family: 'Oswald', sans-serif; font-weight: 600; font-size: 14.5px; color: var(--text); }
        .sb-showcasemeta { font-size: 11.5px; color: var(--text-dim); line-height: 1.4; margin-top: -4px; }
        .sb-addbtn {
          display: flex; align-items: center; justify-content: center; gap: 6px;
          padding: 9px; border-radius: 8px; border: 1px dashed var(--line);
          background: transparent; color: var(--text-dim); font-size: 13px; cursor: pointer;
        }
        .sb-addbtn:hover { border-color: var(--teal); color: var(--teal); }
        .sb-section-gap { margin-top: 18px; }

        .sb-dangerbtn {
          display: flex; align-items: center; justify-content: center; gap: 7px;
          width: 100%; padding: 11px; border-radius: 9px;
          border: 1px solid rgba(220,38,38,0.4); background: rgba(220,38,38,0.08);
          color: var(--danger); font-weight: 600; font-size: 13px; cursor: pointer; margin-top: 8px;
        }

        .sb-footer {
          margin-top: 40px;
          padding: 24px 20px 16px;
          display: flex;
          flex-direction: column;
          align-items: center;
          gap: 10px;
          opacity: 0.5;
          transition: opacity 300ms var(--ease);
        }
        .sb-footer:hover { opacity: 0.8; }
        .sb-footer-logo { width: 80px; height: auto; }
        .sb-footer-text {
          font-family: 'Plus Jakarta Sans', system-ui, sans-serif;
          font-weight: 500;
          font-size: 13px;
          letter-spacing: 0.02em;
          color: var(--text-dim);
        }

        .sb-loginwrap { max-width: 400px; margin: 36px auto 0; }
        .sb-loginroles {
          display: flex; gap: 4px;
          padding: 5px;
          border-radius: 999px;
          background: rgba(255,255,255,0.4);
          backdrop-filter: blur(16px);
          -webkit-backdrop-filter: blur(16px);
          box-shadow: inset 0 1px 0 rgba(255,255,255,0.5), 0 0 0 1px var(--line);
          margin-bottom: 20px;
        }
        .sb-loginrole {
          flex: 1; display: flex; align-items: center; justify-content: center; gap: 7px;
          padding: 11px; border-radius: 999px; border: none;
          background: transparent; color: var(--text-dim);
          font-weight: 600; font-size: 13px; cursor: pointer;
          transition: background 300ms var(--ease), color 200ms var(--ease), transform 300ms var(--ease-spring);
        }
        .sb-loginrole:hover { background: rgba(255,255,255,0.4); }
        .sb-loginrole:active { transform: scale(0.97); }
        .sb-loginrole.active {
          background: rgba(255,255,255,0.8);
          color: var(--text);
          box-shadow: inset 0 1px 0 rgba(255,255,255,0.7), 0 1px 4px rgba(10,10,10,0.08);
        }
        .sb-loginfield-label { font-size: 12px; color: var(--text-dim); margin-bottom: 6px; display: block; font-weight: 500; letter-spacing: 0.01em; }
        .sb-loginerror { color: var(--danger); font-size: 12.5px; margin-top: 12px; text-align: center; font-weight: 500; }
        .sb-loginhint { color: var(--text-dim); font-size: 11.5px; margin-top: 14px; text-align: center; line-height: 1.5; }

        .sb-sharecard {
          margin-top: 14px;
          display: flex;
          flex-direction: column;
          align-items: center;
          padding: 8px;
          border-radius: calc(var(--shell-radius) + 4px);
          background: rgba(255,255,255,0.35);
          box-shadow:
            inset 0 1px 0 rgba(255,255,255,0.5),
            0 0 0 1px var(--line),
            0 12px 36px -18px rgba(10,10,10,0.12);
        }
        .sb-sharecard .sb-label {
          padding-top: 16px;
          font-size: 11px;
          letter-spacing: 0.14em;
        }
        .sb-sharecard svg {
          margin: 12px 0 16px;
          border-radius: 16px !important;
          box-shadow: 0 4px 20px rgba(10,10,10,0.1);
        }
        .sb-qrhint { color: var(--text-dim); font-size: 11px; text-align: center; margin-top: 10px; max-width: 260px; line-height: 1.5; }

        @media (max-width: 480px) {
          .sb-root { padding: 16px 12px 40px; }
          .sb-title { font-size: 22px; }
          .sb-tab span { display: none; }
          .sb-tab { padding: 12px 6px; }
          .sb-header { flex-direction: column; }
          .sb-sessionbox { align-items: center; }
          .sb-loginwrap { margin: 24px auto 0; }
          .sb-card { padding: 14px; }
          .sb-scoreboxes { display: grid; grid-template-columns: repeat(5, 1fr); gap: 7px; }
          .sb-scorebox { min-width: 0; padding: 14px 0; font-size: 15px; }
          .sb-personcard { padding: 14px 12px 12px; }
          .sb-peoplegrid { grid-template-columns: repeat(auto-fill, minmax(130px, 1fr)); gap: 9px; }
          .sb-savebtn { padding: 15px; font-size: 15px; }
          .sb-chip { padding: 10px 13px; }
        }
      `}</style>

      <div className="sb-wrap">
        <div className="sb-header">
          <div className="sb-hero-shell">
            <div className="sb-hero-core">
              <div className="sb-hero-accent" />
              <h1 className="sb-title">
                {session?.role === "judge" && currentJudgeObj ? (
                  <>Hoş Geldiniz,<br /><span style={{ color: "var(--teal)" }}>{currentJudgeObj.name}</span> !</>
                ) : session?.role === "admin" ? (
                  <>Hoş Geldiniz,<br /><span style={{ color: "var(--teal)" }}>Yönetici</span> !</>
                ) : (
                  <>Değerlendirme & Puanlama<br />Platformuna Hoş Geldiniz !</>
                )}
              </h1>
              <div className="sb-sub">Enter Genç Yetenek Programı</div>
              <div className="sb-hero-line">
                <span className="sb-hero-dot" />
                <span className="sb-hero-dot" />
                <span className="sb-hero-dot" />
              </div>
            </div>
          </div>
          {session && (
            <div className="sb-sessionbox">
              <div className="sb-sessionname">
                {session.role === "admin" ? (
                  <>
                    <ShieldCheck size={12} style={{ verticalAlign: -2, marginRight: 4 }} />
                    <b>Yönetici</b>
                  </>
                ) : (
                  <>Giriş: <b>{currentJudgeObj?.name}</b></>
                )}
              </div>
              <button className="sb-logoutbtn" onClick={handleLogout}>
                <LogOut size={13} /> Çıkış Yap
              </button>
            </div>
          )}
        </div>

        {!session && (
          <div className="sb-loginwrap">
            <div className="sb-loginroles">
              <button
                className={`sb-loginrole ${loginRole === "judge" ? "active" : ""}`}
                onClick={() => {
                  setLoginRole("judge");
                  setLoginError("");
                }}
              >
                <User size={15} /> Hakem Girişi
              </button>
              <button
                className={`sb-loginrole ${loginRole === "admin" ? "active" : ""}`}
                onClick={() => {
                  setLoginRole("admin");
                  setLoginError("");
                }}
              >
                <ShieldCheck size={15} /> Yönetici Girişi
              </button>
            </div>

            {loginRole === "judge" ? (
              <div className="sb-card" style={{ padding: 24 }}>
                <label className="sb-loginfield-label">Kullanıcı adı</label>
                <input
                  className="sb-input"
                  style={{ width: "100%", marginBottom: 16 }}
                  value={loginUsername}
                  onChange={(e) => setLoginUsername(e.target.value)}
                  placeholder="Kullanıcı adınızı girin"
                  autoComplete="username"
                />
                <label className="sb-loginfield-label">Şifre</label>
                <input
                  className="sb-input"
                  style={{ width: "100%", marginBottom: 20 }}
                  type="password"
                  value={loginPassword}
                  onChange={(e) => setLoginPassword(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && handleJudgeLogin()}
                  placeholder="Şifrenizi girin"
                  autoComplete="current-password"
                />
                <button className="sb-savebtn" onClick={handleJudgeLogin} disabled={!configLoaded}>
                  <Lock size={15} /> Giriş Yap
                </button>
                {loginError && <div className="sb-loginerror">{loginError}</div>}
              </div>
            ) : (
              <div className="sb-card" style={{ padding: 24 }}>
                <label className="sb-loginfield-label">Yönetici şifresi</label>
                <input
                  className="sb-input"
                  style={{ width: "100%", marginBottom: 20 }}
                  type="password"
                  value={loginAdminPassword}
                  onChange={(e) => setLoginAdminPassword(e.target.value)}
                  onKeyDown={(e) => e.key === "Enter" && handleAdminLogin()}
                  placeholder="Şifrenizi girin"
                  autoComplete="current-password"
                />
                <button className="sb-savebtn" onClick={handleAdminLogin} disabled={!configLoaded}>
                  <ShieldCheck size={15} /> Yönetici Girişi
                </button>
                {loginError && <div className="sb-loginerror">{loginError}</div>}
              </div>
            )}

            {shareUrl ? (
              <div className="sb-card sb-sharecard">
                <div className="sb-label" style={{ textAlign: "center" }}>Mobil Erişim için Okutunuz</div>
                <QRCodeSVG value={shareUrl} size={172} />
              </div>
            ) : (
              pageUrl && (
                <div className="sb-card sb-sharecard">
                  <div className="sb-label" style={{ textAlign: "center" }}>Mobil Erişim için Okutunuz</div>
                  <div className="sb-qrhint" style={{ marginTop: 0 }}>
                    QR kodunu göstermek için önce yönetici olarak giriş yapıp Ayarlar'dan
                    uygulama linkini kaydetmen gerekiyor.
                  </div>
                </div>
              )
            )}
          </div>
        )}

        {session && config.people.length > 0 && (
          <div className="sb-welcome">

            <div className="sb-marquee">
              <div className="sb-marquee-track">
                {[...config.people, ...config.people].map((p, i) => (
                  <span key={i}>{p.name}</span>
                ))}
              </div>
            </div>
          </div>
        )}

        {session && (
          <div className="sb-tabs" role="tablist">
            <button className={`sb-tab ${tab === "score" ? "active" : ""}`} onClick={() => setTab("score")}>
              <ClipboardList size={16} /> <span>Kişi Puanla</span>
            </button>
            {config.presentationCriteria.length > 0 && (
              <button className={`sb-tab ${tab === "presentation" ? "active" : ""}`} onClick={() => setTab("presentation")}>
                <FileText size={16} /> <span>Sunum Puanla</span>
              </button>
            )}
            <button className={`sb-tab ${tab === "dashboard" ? "active" : ""}`} onClick={() => setTab("dashboard")}>
              <BarChart3 size={16} /> <span>Sonuçlar</span>
            </button>
            {config.people.some((p) => p.featured) && (
              <button className={`sb-tab ${tab === "katilanlar" ? "active" : ""}`} onClick={() => setTab("katilanlar")}>
                <PartyPopper size={16} /> <span>Katılanlar</span>
              </button>
            )}
            {session.role === "admin" && (
              <button className={`sb-tab ${tab === "settings" ? "active" : ""}`} onClick={() => setTab("settings")}>
                <Settings size={16} /> <span>Ayarlar</span>
              </button>
            )}
          </div>
        )}

        {session && tab !== "settings" && tab !== "katilanlar" && (
          <div className="sb-timerbar">
            <TimerIcon size={16} color="var(--amber-dim)" />
            <div className="sb-timerdigits">
              {String(Math.floor(timerRemainingSec / 60)).padStart(2, "0")}:
              {String(timerRemainingSec % 60).padStart(2, "0")}
            </div>
            {guidePdf && (
              <a
                href={guidePdf.data}
                target="_blank"
                rel="noopener noreferrer"
                className="sb-guidebtn"
              >
                <FileText size={14} /> Kılavuz
              </a>
            )}
            {session.role === "admin" && (
              <div className="sb-timercontrols">
                <input
                  type="number"
                  min={1}
                  className="sb-select"
                  style={{ width: 56 }}
                  value={Math.round(timer.durationSec / 60)}
                  onChange={(e) => resetTimer((Number(e.target.value) || 1) * 60)}
                  title="Süre (dakika)"
                />
                <span className="sb-timerunit">dk</span>
                {!timer.running ? (
                  <button className="sb-timerbtn" onClick={startTimer} aria-label="Başlat">
                    <Play size={14} />
                  </button>
                ) : (
                  <button className="sb-timerbtn" onClick={pauseTimer} aria-label="Duraklat">
                    <Pause size={14} />
                  </button>
                )}
                <button className="sb-timerbtn" onClick={() => resetTimer()} aria-label="Sıfırla">
                  <RotateCcw size={14} />
                </button>
              </div>
            )}
          </div>
        )}

        {session && tab === "score" && (
          <>
            {session.role === "admin" && (
              <div className="sb-card">
                <div className="sb-label">Hakem olarak puanla (yönetici modu)</div>
                <div className="sb-chiprow">
                  {config.judges.map((j) => (
                    <button
                      key={j.id}
                      className={`sb-chip ${adminActingJudgeId === j.id ? "active" : ""}`}
                      onClick={() => {
                        setAdminActingJudgeId(j.id);
                        setSelectedPerson(null);
                      }}
                    >
                      {j.name}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {currentJudgeId && (
              <div className="sb-card sb-mylockcard">
                <div className="sb-lockrow">
                  <div>
                    <div className="sb-lockstatus">
                      {myLocked ? (
                        <>
                          <Lock size={15} color="var(--danger)" /> Değerlendirmen tamamlandı
                        </>
                      ) : (
                        <>
                          <Check size={15} color="var(--teal)" /> Değerlendirme devam ediyor
                        </>
                      )}
                    </div>
                    <div className="sb-qrhint" style={{ textAlign: "left", margin: "4px 0 0" }}>
                      {myLocked
                        ? "Değişiklik yapmak için kilidi aç."
                        : "Bitirdiğinde kilitle, puanların korunsun."}
                    </div>
                  </div>
                  <button
                    className={myLocked ? "sb-addbtn" : "sb-dangerbtn"}
                    style={{ width: "auto", padding: "10px 16px", marginTop: 0 }}
                    onClick={() => setJudgeLock(currentJudgeId, !myLocked)}
                  >
                    {myLocked ? (
                      <>
                        <Lock size={14} /> Kilidi Aç
                      </>
                    ) : (
                      <>
                        <Lock size={14} /> Değerlendirmemi Tamamla
                      </>
                    )}
                  </button>
                </div>
              </div>
            )}

            {currentJudgeId && (
              <div className="sb-card">
                <div className="sb-label">Değerlendir</div>
                <div className="sb-myprogress">
                  <div className="sb-progresshead">
                    <span>İlerleme</span>
                    <span className="sb-progresspct">
                      {myDoneCount}/{myTotalCount}
                    </span>
                  </div>
                  <div className="sb-progresstrack">
                    <div className="sb-progressfill" style={{ width: `${myPct}%` }} />
                  </div>
                </div>
                <div className="sb-chiprow" style={{ marginBottom: 12 }}>
                  <button
                    className={`sb-chip ${showOnlyUnscored ? "active" : ""}`}
                    onClick={() => setShowOnlyUnscored((v) => !v)}
                  >
                    {showOnlyUnscored ? <Check size={13} style={{ marginRight: 5, verticalAlign: -2 }} /> : null}
                    Değerlendirilmemiş
                  </button>
                </div>
                {groupOptions.length > 1 && (
                  <div className="sb-chiprow" style={{ marginBottom: 12 }}>
                    <button
                      className={`sb-chip ${scoreGroupFilter === "all" ? "active" : ""}`}
                      onClick={() => setScoreGroupFilter("all")}
                    >
                      Tüm gruplar
                    </button>
                    {groupOptions.map((g) => (
                      <button
                        key={g.id}
                        className={`sb-chip ${scoreGroupFilter === g.id ? "active" : ""}`}
                        onClick={() => setScoreGroupFilter(g.id)}
                      >
                        {g.name}
                      </button>
                    ))}
                  </div>
                )}
                {scoreVisiblePeople.length === 0 ? (
                  <div className="sb-empty">🎉 Tüm değerlendirmeler tamamlandı!</div>
                ) : (
                <div className="sb-peoplegrid">
                  {scoreVisiblePeople.map((p) => {
                    const i = p.idx;
                    const done = !!ratings[`r:${currentJudgeId}:${i}`];
                    return (
                      <button
                        key={i}
                        className={`sb-personcard ${selectedPerson === i ? "active" : ""}`}
                        onClick={() => setSelectedPerson(i)}
                      >
                        <div className="ptop">
                          <Avatar photo={p.photo} name={p.name} size={76} shape="square" />
                          <div className="pname">{p.name}</div>
                        </div>
                        {(p.school || p.department) && (
                          <div className="pmeta">
                            {p.school}
                            {p.school && p.department ? " · " : ""}
                            {p.department}
                          </div>
                        )}
                        <div className="pstatus">
                          {done ? (
                            <>
                              <Check size={12} className="check" /> Tamamlandı
                            </>
                          ) : (
                            "Bekliyor"
                          )}
                        </div>
                      </button>
                    );
                  })}
                </div>
                )}
              </div>
            )}

            {currentJudgeId && selectedPerson !== null && (
              <div className="sb-card">
                <div className="sb-label">Değerlendirme</div>
                {config.locked && (
                  <div className="sb-lockbanner">
                    <Lock size={14} /> Puanlama yönetici tarafından kilitlendi — değişiklik yapılamaz.
                  </div>
                )}
                {!config.locked && myLocked && (
                  <div className="sb-lockbanner">
                    <Lock size={14} /> Puanlamanı tamamladın ve kilitledin. Değişiklik yapmak için önce kilidi aç.
                  </div>
                )}
                <div className="sb-profilehead">
                  <Avatar photo={config.people[selectedPerson]?.photo} name={config.people[selectedPerson]?.name} size={56} />
                  <div>
                    <div className="pname">{config.people[selectedPerson]?.name}</div>
                    {(config.people[selectedPerson]?.school || config.people[selectedPerson]?.department) && (
                      <div className="pmeta">
                        {config.people[selectedPerson]?.school}
                        {config.people[selectedPerson]?.school && config.people[selectedPerson]?.department ? " · " : ""}
                        {config.people[selectedPerson]?.department}
                      </div>
                    )}
                  </div>
                </div>
                {config.criteria.map((c) => {
                  const val = draftScores[c.name] ?? 3;
                  return (
                    <div className="sb-slider-row" key={c.name}>
                      <div className="sb-slider-head">
                        <span className="sb-crit-name">
                          {c.name}
                          {c.weight !== 1 && <span className="sb-weightbadge">x{c.weight}</span>}
                        </span>
                        <span className="sb-crit-val" key={val}>
                          {val.toFixed(0)}
                        </span>
                      </div>
                      {c.description && <div className="sb-critdesc">{c.description}</div>}
                      <div className="sb-scoreboxes">
                        {Array.from({ length: 5 }, (_, i) => i + 1).map((n) => (
                          <button
                            key={n}
                            type="button"
                            className={`sb-scorebox ${val === n ? "active" : ""}`}
                            onClick={() => {
                              if (scoringLocked) return;
                              setDraftScores((prev) => ({ ...prev, [c.name]: n }));
                            }}
                            disabled={scoringLocked}
                          >
                            {n}
                          </button>
                        ))}
                      </div>
                    </div>
                  );
                })}
                <div className="sb-slider-row">
                  <div className="sb-slider-head">
                    <span className="sb-crit-name">Yorumunuz</span>
                  </div>
                  <textarea
                    className="sb-notearea"
                    value={draftNote}
                    onChange={(e) => setDraftNote(e.target.value)}
                    placeholder="Bu kişi için kısa bir yorum bırakabilirsin..."
                    rows={3}
                    disabled={scoringLocked}
                  />
                </div>
                <button className="sb-savebtn" onClick={saveRating} disabled={saving || scoringLocked}>
                  {scoringLocked ? "Değerlendirme Tamamlandı" : saving ? "Kaydediliyor…" : "Kaydet"}
                </button>
                {savedFlash && <div className="sb-savedflash">Kaydedildi ✓</div>}
                {saveError && <div className="sb-loginerror">{saveError}</div>}
              </div>
            )}
          </>
        )}

        {session && tab === "presentation" && (
          <>
            {session.role === "admin" && (
              <div className="sb-card">
                <div className="sb-label">Hakem olarak sunum puanla (yönetici modu)</div>
                <div className="sb-chiprow">
                  {config.judges.map((j) => (
                    <button
                      key={j.id}
                      className={`sb-chip ${adminActingJudgeId === j.id ? "active" : ""}`}
                      onClick={() => {
                        setAdminActingJudgeId(j.id);
                        setSelectedPresGroup(null);
                      }}
                    >
                      {j.name}
                    </button>
                  ))}
                </div>
              </div>
            )}

            {currentJudgeId && (
              <div className="sb-card">
                <div className="sb-label">Sunum Grubu Seç</div>
                <div className="sb-chiprow">
                  {config.presentationGroups.map((g) => {
                    const done = !!presRatings[`pr:${currentJudgeId}:${g.id}`];
                    return (
                      <button
                        key={g.id}
                        className={`sb-chip ${selectedPresGroup === g.id ? "active" : ""}`}
                        onClick={() => setSelectedPresGroup(g.id)}
                      >
                        {g.name}
                        {done && <Check size={12} style={{ marginLeft: 5, verticalAlign: -1 }} />}
                      </button>
                    );
                  })}
                </div>
              </div>
            )}

            {currentJudgeId && selectedPresGroup !== null && (
              <div className="sb-card">
                <div className="sb-label">
                  {config.presentationGroups.find((g) => g.id === selectedPresGroup)?.name} — Sunum Değerlendirme
                </div>
                {config.locked && (
                  <div className="sb-lockbanner">
                    <Lock size={14} /> Puanlama yönetici tarafından kilitlendi — değişiklik yapılamaz.
                  </div>
                )}
                {!config.locked && myLocked && (
                  <div className="sb-lockbanner">
                    <Lock size={14} /> Değerlendirmen tamamlandı. Değişiklik yapmak için kilidi aç.
                  </div>
                )}
                {config.presentationCriteria.map((c) => {
                  const val = draftPresScores[c.name] ?? 3;
                  return (
                    <div className="sb-slider-row" key={c.name}>
                      <div className="sb-slider-head">
                        <span className="sb-crit-name">
                          {c.name}
                          {c.weight !== 1 && <span className="sb-weightbadge">x{c.weight}</span>}
                        </span>
                        <span className="sb-crit-val" key={val}>
                          {val.toFixed(0)}
                        </span>
                      </div>
                      {c.description && <div className="sb-critdesc">{c.description}</div>}
                      <div className="sb-scoreboxes">
                        {Array.from({ length: 5 }, (_, i) => i + 1).map((n) => (
                          <button
                            key={n}
                            type="button"
                            className={`sb-scorebox ${val === n ? "active" : ""}`}
                            onClick={() => {
                              if (scoringLocked) return;
                              setDraftPresScores((prev) => ({ ...prev, [c.name]: n }));
                            }}
                            disabled={scoringLocked}
                          >
                            {n}
                          </button>
                        ))}
                      </div>
                    </div>
                  );
                })}
                <div className="sb-slider-row">
                  <div className="sb-slider-head">
                    <span className="sb-crit-name">Yorumunuz</span>
                  </div>
                  <textarea
                    className="sb-notearea"
                    value={draftPresNote}
                    onChange={(e) => setDraftPresNote(e.target.value)}
                    placeholder="Bu sunum için kısa bir yorum bırakabilirsiniz..."
                    rows={3}
                    disabled={scoringLocked}
                  />
                </div>
                <button className="sb-savebtn" onClick={savePresRating} disabled={presSaving || scoringLocked}>
                  {scoringLocked ? "Değerlendirme Tamamlandı" : presSaving ? "Kaydediliyor…" : "Kaydet"}
                </button>
                {presSavedFlash && <div className="sb-savedflash">Kaydedildi ✓</div>}
                {presSaveError && <div className="sb-loginerror">{presSaveError}</div>}
              </div>
            )}
          </>
        )}

        {session && tab === "dashboard" && (
          <>
            <div className="sb-section-divider" style={{ marginTop: 0, paddingTop: 0, borderTop: "none" }}>
              <ClipboardList size={14} /> Kişi Puanlama Sonuçları
            </div>
            <div className="sb-card sb-progresscard">
              <div className="sb-progresshead">
                <span>Tamamlanma Durumu</span>
                <span className="sb-progresspct">%{completionPct}</span>
              </div>
              <div className="sb-progresstrack">
                <div className="sb-progressfill" style={{ width: `${completionPct}%` }} />
              </div>
              <div className="sb-progresssub">
                {totalDoneRatings}/{totalPossibleRatings} değerlendirme tamamlandı
              </div>
            </div>

            <div className="sb-ticker">
              <Radio size={14} color="var(--amber-dim)" />
              <div className="sb-ticker-items">
                {recentActivity.length === 0 && <span>Henüz puanlama yok.</span>}
                {recentActivity.map((r, i) => (
                  <span key={i}>
                    <b>{r.judgeName}</b> → {r.personName} puanladı
                  </span>
                ))}
              </div>
            </div>

            <div className="sb-viewtoggle">
              <button
                className={`sb-viewbtn ${dashView === "chart" ? "active" : ""}`}
                onClick={() => setDashView("chart")}
              >
                <BarChart2 size={14} /> Grafik
              </button>
              <button
                className={`sb-viewbtn ${dashView === "list" ? "active" : ""}`}
                onClick={() => setDashView("list")}
              >
                <List size={14} /> Liste
              </button>
              <button
                className={`sb-viewbtn ${dashView === "compare" ? "active" : ""}`}
                onClick={() => setDashView("compare")}
              >
                <RadarIcon size={14} /> Karşılaştır
              </button>
            </div>

            {dashView === "compare" ? (
              <CompareView
                people={peopleFull}
                criteria={config.criteria}
                personCriteriaAverages={personCriteriaAverages}
                selection={compareSelection}
                setSelection={setCompareSelection}
              />
            ) : (
              <>
                {groupOptions.length > 1 && (
                  <div className="sb-viewtoggle">
                    <button
                      className={`sb-viewbtn ${dashGroupFilter === "all" ? "active" : ""}`}
                      onClick={() => setDashGroupFilter("all")}
                    >
                      <Layers size={14} /> Tüm gruplar
                    </button>
                    {groupOptions.map((g) => (
                      <button
                        key={g.id}
                        className={`sb-viewbtn ${dashGroupFilter === g.id ? "active" : ""}`}
                        onClick={() => setDashGroupFilter(g.id)}
                      >
                        {g.name}
                      </button>
                    ))}
                  </div>
                )}

                {dashGroupFilter === "all" ? (
                  groupBlocks.map((g) => (
                    <div className="sb-card" key={g.id}>
                      {groupOptions.length > 1 && (
                        <div className="sb-groupheader">
                          <span className="gname">{g.name}</span>
                          <span className="gcount">{g.members.length} kişi</span>
                        </div>
                      )}
                      <ScoreBlock items={g.members} view={dashView} />
                    </div>
                  ))
                ) : (
                  <div className="sb-card">
                    <ScoreBlock
                      items={groupBlocks.find((g) => g.id === dashGroupFilter)?.members ?? []}
                      view={dashView}
                    />
                  </div>
                )}
              </>
            )}

            {config.presentationGroups.length > 0 && (
              <>
                <div className="sb-section-divider">
                  <FileText size={14} /> Sunum Sonuçları
                </div>
                <div className="sb-viewtoggle">
                  <button
                    className={`sb-viewbtn ${presDashView === "chart" ? "active" : ""}`}
                    onClick={() => setPresDashView("chart")}
                  >
                    <BarChart2 size={14} /> Grafik
                  </button>
                  <button
                    className={`sb-viewbtn ${presDashView === "list" ? "active" : ""}`}
                    onClick={() => setPresDashView("list")}
                  >
                    <List size={14} /> Liste
                  </button>
                </div>
                <div className="sb-card">
                  <ScoreBlock items={sortByScore(presGroupsFull)} view={presDashView} />
                </div>
              </>
            )}
          </>
        )}

        {session && tab === "katilanlar" && (
          <div className="sb-katilanlar">
            <FireworksCanvas active={tab === "katilanlar"} />
            <div className="sb-katilanlar-title">🎉 Aramıza Katılanlar</div>
            {groupOptions
              .map((g) => ({
                ...g,
                members: config.people.filter(
                  (p) =>
                    p.featured &&
                    (g.id === "__unassigned" ? !knownGroupIds.has(p.groupId) : p.groupId === g.id)
                ),
              }))
              .filter((g) => g.members.length > 0)
              .map((g) => (
                <div className="sb-katilanlar-group" key={g.id}>
                  {groupOptions.length > 1 && <div className="sb-showcase-title">{g.name}</div>}
                  <div className="sb-showcasegrid">
                    {g.members.map((p, i) => (
                      <div className="sb-showcasecard" key={i}>
                        <Avatar photo={p.photo} name={p.name} size={130} shape="square" />
                        <div className="sb-showcasename">{p.name}</div>
                        {(p.school || p.department) && (
                          <div className="sb-showcasemeta">
                            {p.school}
                            {p.school && p.department ? " · " : ""}
                            {p.department}
                          </div>
                        )}
                      </div>
                    ))}
                  </div>
                </div>
              ))}
          </div>
        )}

        {session?.role === "admin" && tab === "settings" && (
          <>
            <div className="sb-card">
              <div className="sb-label">Puanlama Durumu</div>
              <div className="sb-lockrow">
                <div>
                  <div className="sb-lockstatus">
                    {config.locked ? (
                      <>
                        <Lock size={15} color="var(--danger)" /> Kilitli
                      </>
                    ) : (
                      <>
                        <Check size={15} color="var(--teal)" /> Açık — hakemler puan girebilir
                      </>
                    )}
                  </div>
                  <div className="sb-qrhint" style={{ textAlign: "left", margin: "4px 0 0" }}>
                    Kilitlendiğinde hiçbir hakem puan kaydedemez veya değiştiremez.
                  </div>
                </div>
                <button
                  className={config.locked ? "sb-addbtn" : "sb-dangerbtn"}
                  style={{ width: "auto", padding: "10px 16px", marginTop: 0 }}
                  onClick={() => saveConfig({ ...config, locked: !config.locked })}
                >
                  {config.locked ? (
                    <>
                      <Lock size={14} /> Kilidi Aç
                    </>
                  ) : (
                    <>
                      <Lock size={14} /> Puanlamayı Kilitle
                    </>
                  )}
                </button>
              </div>
            </div>

            <div className="sb-card">
              <div className="sb-label">Değerlendirme Kılavuzu (PDF)</div>
              {guidePdf ? (
                <div className="sb-lockrow">
                  <div>
                    <div className="sb-lockstatus">
                      <FileText size={15} color="var(--teal)" /> {guidePdf.name}
                    </div>
                    <div className="sb-qrhint" style={{ textAlign: "left", margin: "4px 0 0" }}>
                      Tüm hakemler bu kılavuzu görebiliyor.
                    </div>
                  </div>
                  <button
                    className="sb-dangerbtn"
                    style={{ width: "auto", padding: "10px 16px", marginTop: 0 }}
                    onClick={removeGuidePdf}
                  >
                    <Trash2 size={14} /> Kaldır
                  </button>
                </div>
              ) : (
                <>
                  <div className="sb-qrhint" style={{ textAlign: "left", margin: "0 0 10px" }}>
                    Yüklediğin PDF, tüm hakemlerin ekranında kılavuz olarak görünür ve açılabilir.
                  </div>
                  <label className="sb-savebtn" style={{ cursor: "pointer" }}>
                    <Upload size={15} /> {guideUploading ? "Yükleniyor…" : "PDF Yükle"}
                    <input
                      type="file"
                      accept="application/pdf"
                      style={{ display: "none" }}
                      onChange={(e) => {
                        const file = e.target.files?.[0];
                        uploadGuidePdf(file);
                        e.target.value = "";
                      }}
                      disabled={guideUploading}
                    />
                  </label>
                </>
              )}
              {guideError && <div className="sb-loginerror">{guideError}</div>}
            </div>

            {config.people.some((p) => p.featured) && (
              <div className="sb-card">
                <div className="sb-showcase-title">Aramıza Katılanlar</div>
                <div className="sb-showcasegrid">
                  {config.people
                    .filter((p) => p.featured)
                    .map((p, i) => (
                      <div className="sb-showcasecard" key={i}>
                        <Avatar photo={p.photo} name={p.name} size={130} shape="square" />
                        <div className="sb-showcasename">{p.name}</div>
                        {(p.school || p.department) && (
                          <div className="sb-showcasemeta">
                            {p.school}
                            {p.school && p.department ? " · " : ""}
                            {p.department}
                          </div>
                        )}
                      </div>
                    ))}
                </div>
              </div>
            )}

            <div className="sb-card">
              <div className="sb-label">Hakemler (Kullanıcı Adı / Şifre)</div>
              <div className="sb-settings-list">
                {draftConfig.judges.map((j, i) => (
                  <div className="sb-settings-item" key={j.id}>
                    <input
                      className="sb-input"
                      value={j.name}
                      placeholder="İsim"
                      onChange={(e) => {
                        const next = { ...draftConfig, judges: [...draftConfig.judges] };
                        next.judges[i] = { ...j, name: e.target.value };
                        setDraftConfig(next);
                      }}
                      onBlur={() => saveConfig(draftConfig)}
                    />
                    <input
                      className="sb-input"
                      value={j.username}
                      placeholder="Kullanıcı adı"
                      onChange={(e) => {
                        const next = { ...draftConfig, judges: [...draftConfig.judges] };
                        next.judges[i] = { ...j, username: e.target.value };
                        setDraftConfig(next);
                      }}
                      onBlur={() => saveConfig(draftConfig)}
                    />
                    <input
                      className="sb-input"
                      value={j.password}
                      placeholder="Şifre"
                      onChange={(e) => {
                        const next = { ...draftConfig, judges: [...draftConfig.judges] };
                        next.judges[i] = { ...j, password: e.target.value };
                        setDraftConfig(next);
                      }}
                      onBlur={() => saveConfig(draftConfig)}
                    />
                    <button
                      className="sb-iconbtn"
                      onClick={() => {
                        const next = { ...draftConfig, judges: draftConfig.judges.filter((_, idx) => idx !== i) };
                        saveConfig(next);
                      }}
                      aria-label="Hakemi sil"
                    >
                      <Trash2 size={15} />
                    </button>
                    <button
                      className={`sb-iconbtn ${judgeLocks[j.id] ? "sb-iconbtn-active" : ""}`}
                      onClick={() => setJudgeLock(j.id, !judgeLocks[j.id])}
                      aria-label={judgeLocks[j.id] ? "Hakemin kilidini aç" : "Hakem kilitli değil"}
                      title={judgeLocks[j.id] ? "Kilitli — açmak için tıkla" : "Açık"}
                    >
                      <Lock size={15} />
                    </button>
                  </div>
                ))}
                <button
                  className="sb-addbtn"
                  onClick={() => {
                    const id = newJudgeId(draftConfig.judges);
                    const n = draftConfig.judges.length + 1;
                    saveConfig({
                      ...draftConfig,
                      judges: [
                        ...draftConfig.judges,
                        { id, name: `Hakem ${n}`, username: `hakem${n}`, password: "1234" },
                      ],
                    });
                  }}
                >
                  <Plus size={14} /> Hakem ekle
                </button>
              </div>
            </div>

            <div className="sb-card">
              <div className="sb-label">Gruplar</div>
              <div className="sb-settings-list">
                {draftConfig.groups.map((g, i) => (
                  <div className="sb-settings-item" key={g.id}>
                    <input
                      className="sb-input"
                      value={g.name}
                      onChange={(e) => {
                        const next = { ...draftConfig, groups: [...draftConfig.groups] };
                        next.groups[i] = { ...g, name: e.target.value };
                        setDraftConfig(next);
                      }}
                      onBlur={() => saveConfig(draftConfig)}
                    />
                    <button
                      className="sb-iconbtn"
                      onClick={() => {
                        if (draftConfig.groups.length <= 1) return;
                        const remaining = draftConfig.groups.filter((x) => x.id !== g.id);
                        const fallbackId = remaining[0].id;
                        const nextPeople = draftConfig.people.map((p) =>
                          p.groupId === g.id ? { ...p, groupId: fallbackId } : p
                        );
                        saveConfig({ ...draftConfig, groups: remaining, people: nextPeople });
                      }}
                      aria-label="Grubu sil"
                    >
                      <Trash2 size={15} />
                    </button>
                  </div>
                ))}
                <button
                  className="sb-addbtn"
                  onClick={() => {
                    const id = newGroupId(draftConfig.groups);
                    saveConfig({
                      ...draftConfig,
                      groups: [...draftConfig.groups, { id, name: `Grup ${draftConfig.groups.length + 1}` }],
                    });
                  }}
                >
                  <Plus size={14} /> Grup ekle
                </button>
              </div>
            </div>

            <div className="sb-card">
              <div className="sb-label">Sunum Grupları</div>
              <div className="sb-qrhint" style={{ textAlign: "left", margin: "0 0 10px" }}>
                Yarışmacı gruplarından bağımsızdır — sunumları nasıl ayırmak istersen (örn. Elektrik, Veri) burada kendi grubunu oluşturabilirsin.
              </div>
              <div className="sb-settings-list">
                {draftConfig.presentationGroups.map((g, i) => (
                  <div className="sb-settings-item" key={g.id}>
                    <input
                      className="sb-input"
                      value={g.name}
                      onChange={(e) => {
                        const next = { ...draftConfig, presentationGroups: [...draftConfig.presentationGroups] };
                        next.presentationGroups[i] = { ...g, name: e.target.value };
                        setDraftConfig(next);
                      }}
                      onBlur={() => saveConfig(draftConfig)}
                    />
                    <button
                      className="sb-iconbtn"
                      onClick={() => {
                        if (draftConfig.presentationGroups.length <= 1) return;
                        const remaining = draftConfig.presentationGroups.filter((x) => x.id !== g.id);
                        saveConfig({ ...draftConfig, presentationGroups: remaining });
                      }}
                      aria-label="Sunum grubunu sil"
                    >
                      <Trash2 size={15} />
                    </button>
                  </div>
                ))}
                <button
                  className="sb-addbtn"
                  onClick={() => {
                    const id = newPresGroupId(draftConfig.presentationGroups);
                    saveConfig({
                      ...draftConfig,
                      presentationGroups: [...draftConfig.presentationGroups, { id, name: `Sunum Grubu ${draftConfig.presentationGroups.length + 1}` }],
                    });
                  }}
                >
                  <Plus size={14} /> Sunum grubu ekle
                </button>
              </div>
            </div>

            <div className="sb-card">
              <div className="sb-label">Yarışmacılar</div>
              <div className="sb-settings-list">
                {draftConfig.people.map((p, i) => (
                  <div className="sb-settings-item" key={i} style={{ flexDirection: "column", alignItems: "stretch", border: "1px solid var(--line)", borderRadius: 10, padding: 10 }}>
                    <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
                      <div className="sb-photoupload">
                        <label>
                          {p.photo ? <Avatar photo={p.photo} name={p.name} size={84} shape="square" /> : <Plus size={26} />}
                          <input
                            type="file"
                            accept="image/*"
                            onChange={async (e) => {
                              const file = e.target.files?.[0];
                              if (!file) return;
                              try {
                                const dataUrl = await resizeImageFile(file);
                                const next = { ...draftConfig, people: [...draftConfig.people] };
                                next.people[i] = { ...draftConfig.people[i], photo: dataUrl };
                                saveConfig(next);
                              } catch (err) {}
                              e.target.value = "";
                            }}
                          />
                        </label>
                      </div>
                      <input
                        className="sb-input"
                        placeholder="İsim"
                        value={p.name}
                        onChange={(e) => {
                          const next = { ...draftConfig, people: [...draftConfig.people] };
                          next.people[i] = { ...p, name: e.target.value };
                          setDraftConfig(next);
                        }}
                        onBlur={() => saveConfig(draftConfig)}
                      />
                      <select
                        className="sb-select"
                        value={p.groupId}
                        onChange={(e) => {
                          const next = { ...draftConfig, people: [...draftConfig.people] };
                          next.people[i] = { ...p, groupId: e.target.value };
                          saveConfig(next);
                        }}
                      >
                        {draftConfig.groups.map((g) => (
                          <option key={g.id} value={g.id}>
                            {g.name}
                          </option>
                        ))}
                      </select>
                      <button
                        className={`sb-iconbtn ${p.featured ? "sb-iconbtn-active" : ""}`}
                        onClick={() => {
                          const next = { ...draftConfig, people: [...draftConfig.people] };
                          next.people[i] = { ...p, featured: !p.featured };
                          saveConfig(next);
                        }}
                        aria-label="Aramıza Katılanlar'da göster"
                        title="Aramıza Katılanlar'da göster"
                      >
                        <Star size={15} fill={p.featured ? "currentColor" : "none"} />
                      </button>
                      <button
                        className="sb-iconbtn"
                        onClick={() => {
                          const next = { ...draftConfig, people: draftConfig.people.filter((_, idx) => idx !== i) };
                          saveConfig(next);
                        }}
                        aria-label="Yarışmacıyı sil"
                      >
                        <Trash2 size={15} />
                      </button>
                    </div>
                    <div style={{ display: "flex", gap: 8, marginTop: 8 }}>
                      <input
                        className="sb-input"
                        placeholder="Okul"
                        value={p.school || ""}
                        onChange={(e) => {
                          const next = { ...draftConfig, people: [...draftConfig.people] };
                          next.people[i] = { ...p, school: e.target.value };
                          setDraftConfig(next);
                        }}
                        onBlur={() => saveConfig(draftConfig)}
                      />
                      <input
                        className="sb-input"
                        placeholder="Bölüm"
                        value={p.department || ""}
                        onChange={(e) => {
                          const next = { ...draftConfig, people: [...draftConfig.people] };
                          next.people[i] = { ...p, department: e.target.value };
                          setDraftConfig(next);
                        }}
                        onBlur={() => saveConfig(draftConfig)}
                      />
                    </div>
                  </div>
                ))}
                <button
                  className="sb-addbtn"
                  onClick={() =>
                    saveConfig({
                      ...draftConfig,
                      people: [
                        ...draftConfig.people,
                        { name: `Kişi ${draftConfig.people.length + 1}`, groupId: draftConfig.groups[0].id, photo: "", school: "", department: "", featured: false },
                      ],
                    })
                  }
                >
                  <Plus size={14} /> Yarışmacı ekle
                </button>
              </div>
            </div>

            <div className="sb-card">
              <div className="sb-label">Kriterler</div>
              <div className="sb-settings-list">
                {draftConfig.criteria.map((c, i) => (
                  <div key={i} style={{ border: "1px solid var(--line)", borderRadius: 10, padding: 10 }}>
                    <div className="sb-settings-item" style={{ padding: 0, border: "none" }}>
                      <input
                        className="sb-input"
                        value={c.name}
                        placeholder="Kriter adı"
                        onChange={(e) => {
                          const next = { ...draftConfig, criteria: [...draftConfig.criteria] };
                          next.criteria[i] = { ...c, name: e.target.value };
                          setDraftConfig(next);
                        }}
                        onBlur={() => saveConfig(draftConfig)}
                      />
                      <input
                        className="sb-select"
                        type="number"
                        min={0.5}
                        max={5}
                        step={0.5}
                        style={{ width: 64, textAlign: "center" }}
                        value={c.weight}
                        title="Ağırlık çarpanı"
                        onChange={(e) => {
                          const next = { ...draftConfig, criteria: [...draftConfig.criteria] };
                          next.criteria[i] = { ...c, weight: Number(e.target.value) || 1 };
                          setDraftConfig(next);
                        }}
                        onBlur={() => saveConfig(draftConfig)}
                      />
                      <button
                        className="sb-iconbtn"
                        onClick={() => {
                          const next = { ...draftConfig, criteria: draftConfig.criteria.filter((_, idx) => idx !== i) };
                          saveConfig(next);
                        }}
                        aria-label="Kriteri sil"
                      >
                        <Trash2 size={15} />
                      </button>
                    </div>
                    <textarea
                      className="sb-notearea"
                      style={{ marginTop: 8, minHeight: 44, fontSize: 12.5 }}
                      value={c.description || ""}
                      placeholder="Bu kriter için açıklama (hakem puan verirken görecek, opsiyonel)"
                      rows={2}
                      onChange={(e) => {
                        const next = { ...draftConfig, criteria: [...draftConfig.criteria] };
                        next.criteria[i] = { ...c, description: e.target.value };
                        setDraftConfig(next);
                      }}
                      onBlur={() => saveConfig(draftConfig)}
                    />
                  </div>
                ))}
                <div className="sb-qrhint" style={{ textAlign: "left", margin: "2px 0 4px" }}>
                  Sağdaki sayı ağırlık çarpanıdır — 1 normal, 2 iki kat, 0.5 yarı ağırlık gibi. Genel puan bu ağırlıklara göre hesaplanır.
                </div>
                <button
                  className="sb-addbtn"
                  onClick={() =>
                    saveConfig({ ...draftConfig, criteria: [...draftConfig.criteria, { name: `Kriter ${draftConfig.criteria.length + 1}`, weight: 1, description: "" }] })
                  }
                >
                  <Plus size={14} /> Kriter ekle
                </button>
              </div>
            </div>

            <div className="sb-card">
              <div className="sb-label">Sunum Kriterleri</div>
              <div className="sb-settings-list">
                {draftConfig.presentationCriteria.map((c, i) => (
                  <div key={i} style={{ border: "1px solid var(--line)", borderRadius: 10, padding: 10 }}>
                    <div className="sb-settings-item" style={{ padding: 0, border: "none" }}>
                      <input
                        className="sb-input"
                        value={c.name}
                        placeholder="Kriter adı"
                        onChange={(e) => {
                          const next = { ...draftConfig, presentationCriteria: [...draftConfig.presentationCriteria] };
                          next.presentationCriteria[i] = { ...c, name: e.target.value };
                          setDraftConfig(next);
                        }}
                        onBlur={() => saveConfig(draftConfig)}
                      />
                      <input
                        className="sb-select"
                        type="number"
                        min={0.5}
                        max={5}
                        step={0.5}
                        style={{ width: 64, textAlign: "center" }}
                        value={c.weight}
                        title="Ağırlık çarpanı"
                        onChange={(e) => {
                          const next = { ...draftConfig, presentationCriteria: [...draftConfig.presentationCriteria] };
                          next.presentationCriteria[i] = { ...c, weight: Number(e.target.value) || 1 };
                          setDraftConfig(next);
                        }}
                        onBlur={() => saveConfig(draftConfig)}
                      />
                      <button
                        className="sb-iconbtn"
                        onClick={() => {
                          const next = { ...draftConfig, presentationCriteria: draftConfig.presentationCriteria.filter((_, idx) => idx !== i) };
                          saveConfig(next);
                        }}
                        aria-label="Sunum kriterini sil"
                      >
                        <Trash2 size={15} />
                      </button>
                    </div>
                    <textarea
                      className="sb-notearea"
                      style={{ marginTop: 8, minHeight: 44, fontSize: 12.5 }}
                      value={c.description || ""}
                      placeholder="Bu kriter için açıklama (hakem puan verirken görecek, opsiyonel)"
                      rows={2}
                      onChange={(e) => {
                        const next = { ...draftConfig, presentationCriteria: [...draftConfig.presentationCriteria] };
                        next.presentationCriteria[i] = { ...c, description: e.target.value };
                        setDraftConfig(next);
                      }}
                      onBlur={() => saveConfig(draftConfig)}
                    />
                  </div>
                ))}
                <button
                  className="sb-addbtn"
                  onClick={() =>
                    saveConfig({ ...draftConfig, presentationCriteria: [...draftConfig.presentationCriteria, { name: `Sunum Kriteri ${draftConfig.presentationCriteria.length + 1}`, weight: 1, description: "" }] })
                  }
                >
                  <Plus size={14} /> Sunum kriteri ekle
                </button>
              </div>
            </div>

            <div className="sb-card">
              <div className="sb-label">Yönetici Şifresi</div>
              <input
                className="sb-input"
                style={{ width: "100%" }}
                value={draftConfig.adminPassword}
                onChange={(e) => setDraftConfig({ ...draftConfig, adminPassword: e.target.value })}
                onBlur={() => saveConfig(draftConfig)}
              />
            </div>

            <div className="sb-card">
              <div className="sb-label">Uygulama Linki (QR / Paylaşım için)</div>
              <input
                className="sb-input"
                style={{ width: "100%" }}
                placeholder="https://claude.ai/... (Publish sonrası aldığın link)"
                value={draftConfig.appUrl}
                onChange={(e) => setDraftConfig({ ...draftConfig, appUrl: e.target.value })}
                onBlur={() => saveConfig(draftConfig)}
              />
              <div className="sb-qrhint" style={{ margin: "8px 0 0", textAlign: "left" }}>
                Bu artifact'ı "Share" ile yayınladıktan sonra aldığın genel linki buraya yapıştır —
                giriş ekranındaki QR kod ve link kutusu bundan sonra bu linki gösterir.
              </div>
            </div>

            <div className="sb-card">
              <div className="sb-label">Hakem Notları</div>
              {(() => {
                const groups = {};
                Object.entries(ratings).forEach(([k, r]) => {
                  if (!r || !r.note || !r.note.trim()) return;
                  const parts = k.split(":");
                  const judge = config.judges.find((j) => j.id === parts[1]);
                  const pIdx = Number(parts[2]);
                  const person = config.people[pIdx];
                  const personName = person?.name ?? r.personName ?? "?";
                  if (!groups[personName]) groups[personName] = { photo: person?.photo, notes: [] };
                  groups[personName].notes.push({
                    judgeName: judge?.name ?? r.judgeName ?? "?",
                    note: r.note,
                    ts: r.ts,
                  });
                });
                const personGroups = Object.entries(groups)
                  .map(([personName, g]) => ({
                    personName,
                    photo: g.photo,
                    notes: g.notes.sort((a, b) => b.ts - a.ts),
                    latestTs: Math.max(...g.notes.map((n) => n.ts || 0)),
                  }))
                  .sort((a, b) => b.latestTs - a.latestTs);
                if (personGroups.length === 0) return <div className="sb-empty">Henüz not bırakılmadı.</div>;
                return (
                  <div className="sb-notelist">
                    {personGroups.map((g) => (
                      <div className="sb-notegroup" key={g.personName}>
                        <div className="sb-notegroup-head">
                          <Avatar photo={g.photo} name={g.personName} size={26} />
                          <span>{g.personName}</span>
                        </div>
                        {g.notes.map((n, i) => (
                          <div className="sb-noteitem" key={i}>
                            <div className="sb-noteitem-head">
                              <b>{n.judgeName}</b>
                            </div>
                            <div className="sb-noteitem-text">{n.note}</div>
                          </div>
                        ))}
                      </div>
                    ))}
                  </div>
                );
              })()}
            </div>

            <div className="sb-card">
              <div className="sb-label">Dışa Aktar</div>
              <button className="sb-savebtn" onClick={exportToExcel}>
                <Download size={15} /> Excel olarak indir (.xlsx)
              </button>
              {exportError && <div className="sb-loginerror">{exportError}</div>}
              <button className="sb-savebtn" style={{ marginTop: 10, background: "#e63946" }} onClick={() => setShowReport(true)}>
                <FileText size={15} /> PDF Kapanış Raporu Oluştur
              </button>
            </div>

            <div className="sb-card">
              <div className="sb-label">Tehlikeli Bölge</div>
              <button className="sb-dangerbtn" onClick={clearAllRatings}>
                <RotateCcw size={15} /> Tüm puanları sıfırla
              </button>
            </div>
          </>
        )}

        {session && (
          <div className="sb-footer">
            <img src={LOGO2_SRC} alt="Logo" className="sb-footer-logo" />
            <div className="sb-footer-text">Hoş Geldiniz</div>
          </div>
        )}
      </div>
    </div>
  );
}
