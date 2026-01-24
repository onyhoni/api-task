function generateRandomString(length) {
    const characters = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
    let randomString = '';
    for (let i = 0; i < length; i++) {
        randomString += characters.charAt(
            Math.floor(Math.random() * characters.length)
        );
    }
    return randomString;
}

function generateUniqueCode() {
    // Format CT08
    let code = 'CT08';

    // Random String
    const randomString = generateRandomString(6);
    code += randomString;

    // Tanggal + Bulan + Tahun
    const currentDate = new Date();
    const monthAbbreviations = [
        'JAN',
        'FEB',
        'MAR',
        'APR',
        'MEI',
        'JUN',
        'JUL',
        'AGT',
        'SEP',
        'OKT',
        'NOV',
        'DES',
    ];
    const datePart =
        currentDate.getDate().toString().padStart(2, '0') +
        monthAbbreviations[currentDate.getMonth()] +
        currentDate.getFullYear();
    code += datePart;

    // Nomor Urut (bisa di-generate sesuai kebutuhan)
    // Misalnya, kita tambahkan nomor urut secara sederhana
    // Jika butuh lebih kompleks, sesuaikan dengan kebutuhan sistem Anda
    const noUrut = Math.floor(Math.random() * 10000)
        .toString()
        .padStart(4, '0');
    code += noUrut;

    return code;
}

// Contoh penggunaan
const uniqueCode = generateUniqueCode();
console.log(uniqueCode);
