import { initializeApp } from 'firebase/app';
import { getFirestore, doc, getDocFromServer } from 'firebase/firestore';

const firebaseConfig = {
  projectId: "secure-diagram-hnzsc",
  appId: "1:937146532987:web:14ff8b9039d766963dcede",
  apiKey: "AIzaSyAL5CCYirOmc2zNKvvmAtylNsQOA_FDXfA",
  authDomain: "diagram-hnzsc.firebaseapp.com",
  firestoreDatabaseId: "ai-studio-d21e92a7-5773-41c8-8bcc-034fb2f2d87f",
  storageBucket: "secure-diagram-hnzsc.firebasestorage.app",
  messagingSenderId: "937146532987"
};

const app = initializeApp(firebaseConfig);
export const db = getFirestore(app, firebaseConfig.firestoreDatabaseId);

// Validate connection to Firestore on initialization
async function testConnection() {
  try {
    await getDocFromServer(doc(db, 'test', 'connection'));
    console.log("Connected to Firebase Firestore successfully!");
  } catch (error) {
    if (error instanceof Error && error.message.includes('the client is offline')) {
      console.error("Please check your Firebase configuration or network.");
    } else {
      console.log("Firestore initialized:", error);
    }
  }
}
testConnection();
