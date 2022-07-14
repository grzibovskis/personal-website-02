
/*========================== BULB ======================*/

const light = document.getElementById("bulb");
setInterval(function() {

        if (light.style.display === 'none'){
            light.style.display = 'block';
        } else {
            light.style.display = 'none'
        }

}, 1000);

/*======================= Typewriter Header Section and Button functions========================= */
const actionBtn = document.querySelector('#circle')
const typeWriter = document.querySelector('.typewriter')
const actionBtnSec = document.querySelector('.action_btn_sec')
const aboutMe = document.querySelector('.about_me')
const actionBtnThird = document.querySelector('.action_btn_third')
const knowledge = document.querySelector('.knowledge')
actionBtn.addEventListener('click', () => {
        typeWriter.style.display = 'block';
    
    setInterval(function() { actionBtnSec.style.opacity = '1'; }, 5500);
})
actionBtnSec.addEventListener('click', () => {
    aboutMe.style.display = 'block';
})
actionBtnThird.addEventListener('click', () => {
    knowledge.style.display = 'block';
})





/*========================= button ========================*/
const circle = document.querySelector('#circle');

circle.addEventListener('mouseenter', () => {
    if(!circle.classList.contains('hover')) {
        circle.classList.add('hover');
    }
    
});
circle.addEventListener('mouseleave', () => {
    if(circle.classList.contains('hover')) {
        circle.classList.remove('hover');
    }
});
const circle_sec = document.querySelector('#circle_sec');

circle_sec.addEventListener('mouseenter', () => {
    if(!circle_sec.classList.contains('hover')) {
        circle_sec.classList.add('hover');
    }
    
});
circle_sec.addEventListener('mouseleave', () => {
    if(circle_sec.classList.contains('hover')) {
        circle_sec.classList.remove('hover');
    }
});