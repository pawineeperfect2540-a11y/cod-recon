// ====================================================================
// ตั้งค่า Firebase สำหรับเก็บประวัติย้อนหลัง
// วิธีตั้งค่า: ดูคำแนะนำในแท็บ "ตั้งค่า" ของเว็บ หรืออ่านใน README.md
//
// 1. เข้า https://console.firebase.google.com
// 2. สร้างโปรเจกต์ใหม่ (หรือใช้โปรเจกต์เดียวกับ deposit-form-app เดิมก็ได้
//    เพื่อไม่ต้องตั้งค่าใหม่ทั้งหมด — แค่สร้าง Firestore collection ใหม่ชื่อ
//    "cod_periods" ในโปรเจกต์เดิม)
// 3. ไปที่ Project Settings (รูปเฟือง) > General > Your apps > Web app
//    คัดลอกค่า firebaseConfig มาแทนที่ด้านล่างนี้ทั้งหมด
// 4. เปิดใช้งาน Firestore Database (Build > Firestore Database > Create database)
// ====================================================================

window.FIREBASE_CONFIG = {
  apiKey: "AIzaSyD3xP2lbqerKjzXeAaOCbKUJ8ospHCRSlw",
  authDomain: "perfectads-deposit.firebaseapp.com",
  projectId: "perfectads-deposit",
  storageBucket: "perfectads-deposit.firebasestorage.app",
  messagingSenderId: "825880055779",
  appId: "1:825880055779:web:9c1d3dd57533bacbd69ad6",
  measurementId: "G-95LTQMEE9T"
};

// ตั้งเป็น true เมื่อกรอกค่าด้านบนครบแล้ว เพื่อเปิดใช้งานการบันทึกประวัติ
window.FIREBASE_ENABLED = true;
