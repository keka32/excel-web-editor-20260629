// sync.js — Firebase Firestore sync for multi-device collaboration
// Free tier: 50K reads/day, 20K writes/day, 1GB storage — plenty for Excel files.
// Uses anonymous auth — no login required.

var CloudSync = (function () {
  'use strict';

  var firebaseConfig = {
    apiKey: "AIzaSyDajLamE0J5dpjclHMX9zm1i59rwOL6a28",
    authDomain: "excel-editor32.firebaseapp.com",
    projectId: "excel-editor32",
    storageBucket: "excel-editor32.firebasestorage.app",
    messagingSenderId: "713322690524",
    appId: "1:713322690524:web:ce16f604e22c3c31388119",
  };

  var db = null;
  var ready = false;
  var onRemoteChange = null;
  var lastSync = {};

  function init(cfg) {
    if (cfg) firebaseConfig = cfg;
    try {
      firebase.initializeApp(firebaseConfig);
      firebase.auth().signInAnonymously().then(function () {
        db = firebase.firestore();
        db.enablePersistence({ synchronizeTabs: true }).catch(function () {});
        ready = true;
        console.log('[Sync] Firebase ready (anonymous)');
      }).catch(function (e) {
        console.warn('[Sync] Auth failed:', e.message);
      });
    } catch (e) {
      console.warn('[Sync] Firebase init failed:', e.message);
    }
  }

  function isReady() { return ready && db; }

  function collectionName() { return 'excel-files'; }

  function saveFile(name, dataArray) {
    if (!isReady()) return Promise.resolve(false);
    var docRef = db.collection(collectionName()).doc(encodeURIComponent(name));
    return docRef.set({
      name: name,
      data: dataArray,
      updatedAt: firebase.firestore.FieldValue.serverTimestamp(),
      updatedBy: 'anon',
    }).then(function () {
      lastSync[name] = Date.now();
      return true;
    }).catch(function (e) {
      console.warn('[Sync] Save failed:', name, e.message);
      return false;
    });
  }

  function loadFile(name) {
    if (!isReady()) return Promise.resolve(null);
    return db.collection(collectionName()).doc(encodeURIComponent(name)).get()
      .then(function (doc) {
        if (!doc.exists) return null;
        var d = doc.data();
        lastSync[name] = Date.now();
        return d.data;
      }).catch(function (e) {
        console.warn('[Sync] Load failed:', name, e.message);
        return null;
      });
  }

  function loadAllFiles() {
    if (!isReady()) return Promise.resolve([]);
    return db.collection(collectionName()).orderBy('updatedAt', 'desc').limit(20).get()
      .then(function (snap) {
        var files = [];
        snap.forEach(function (doc) {
          var d = doc.data();
          lastSync[d.name] = Date.now();
          files.push({ name: d.name, data: d.data });
        });
        return files;
      }).catch(function (e) {
        console.warn('[Sync] List failed:', e.message);
        return [];
      });
  }

  function watchChanges(callback) {
    if (!isReady()) return;
    onRemoteChange = callback;
    db.collection(collectionName()).onSnapshot(function (snap) {
      snap.docChanges().forEach(function (change) {
        var d = change.doc.data();
        if (!d || !d.name) return;
        if (lastSync[d.name] && (Date.now() - lastSync[d.name] < 3000)) return; // ignore own saves
        if (change.type === 'removed') return;
        if (onRemoteChange) onRemoteChange(d.name, d.data);
      });
    }, function (e) {
      console.warn('[Sync] Watch error:', e.message);
    });
  }

  function deleteFile(name) {
    if (!isReady()) return Promise.resolve(false);
    return db.collection(collectionName()).doc(encodeURIComponent(name)).delete()
      .then(function () { return true; })
      .catch(function () { return false; });
  }

  return {
    init: init,
    isReady: isReady,
    saveFile: saveFile,
    loadFile: loadFile,
    loadAllFiles: loadAllFiles,
    watchChanges: watchChanges,
    deleteFile: deleteFile,
  };
})();
