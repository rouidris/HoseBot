// [1, 2, 3, 4, 5] --> [-1, -2, -3, -4, -5]
// [1, -2, 3, -4, 5] --> [-1, 2, -3, 4, -5]
// [] --> []

function invert(array) {
    let invertArr = [];
    for(let i = 0; i < array.length; i++){
        invertArr.push(-array[i]);
    }
    return invertArr, array;
}

console.log(invert([1, 2, 3, 4, 5, -6]));