import { initializeApp } from "https://www.gstatic.com/firebasejs/11.0.1/firebase-app.js";
import { getFirestore, initializeFirestore } from "https://www.gstatic.com/firebasejs/11.0.1/firebase-firestore.js";
import { getAnalytics } from "https://www.gstatic.com/firebasejs/11.0.1/firebase-analytics.js";
import { getAuth, GoogleAuthProvider } from "https://www.gstatic.com/firebasejs/11.0.1/firebase-auth.js";

// Your web app's Firebase configuration
const firebaseConfig = {
    apiKey: "AIzaSyAwrYUYcQ6KggiA2Y0ublcA0v0-BpMp4Sw",
    authDomain: "sbe-4d7b4.firebaseapp.com",
    projectId: "sbe-4d7b4",
    storageBucket: "sbe-4d7b4.firebasestorage.app",
    messagingSenderId: "178352953312",
    appId: "1:178352953312:web:0f825092cf70362826fd85",
    measurementId: "G-Q2XSZWGE7Q"
};

// Initialize Firebase
const app = initializeApp(firebaseConfig);

// Using initializeFirestore with experimentalForceLongPolling:true 
// AND specifying the named database "sbedsa"
const db = initializeFirestore(app, {
    experimentalForceLongPolling: true
}, "sbedsa");

const auth = getAuth(app);
const googleProvider = new GoogleAuthProvider();
const analytics = getAnalytics(app);

export { db, auth, googleProvider, analytics };
