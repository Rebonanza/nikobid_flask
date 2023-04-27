const angkaPangkat = document.getElementById('angka-pangkat') 
const pangkat =  document.getElementById('pangkat')
const angkaFaktorial =  document.getElementById('angka-faktorial')
const hasilFaktorial = document.getElementById('hasil-faktor')
const hasilPangkat = document.getElementById('hasil-pangkat')

const btnPangkat = document.getElementById('hitung-pangkat')
const btnFaktor = document.getElementById('hitung-faktorial')

try {
  const data = file.readFileSync('data.txt', 'utf8');
  console.log(data);
} catch (err) {
  console.error(err);
}


function perpangkatan(angka, pangkat){
    return Math.pow(angka,pangkat)
}

function faktorial(angka){
    let hasil = 1
    for(let i=angka; i>0; i--){
        hasil *=i;
    }
    return hasil
}

if(btnPangkat != null){
    btnPangkat.addEventListener('click', ()=>{
        let angka =  angkaPangkat.value
        let powe = pangkat.value
        console.log(angka)
        console.log(powe)
        let hasil = perpangkatan(angka,powe)
        console.log(hasil)
        hasilPangkat.innerText = hasil
            try {
                file.writeFileSync('data.txt', hasil);
                console.log("sudah write")
                // file written successfully
                } catch (err) {
                console.error(err);
                }
           })
}

if(btnFaktor != null){
    btnFaktor.addEventListener('click',()=>{
        let angka =  angkaFaktorial.value
        let hasil = faktorial(angka)
        hasilFaktorial.innerText = hasil
    })
}
