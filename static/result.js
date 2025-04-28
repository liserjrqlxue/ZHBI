function pollStatus(taskID) {
    fetch('/status?taskID=' + taskID)
        .then(response => response.json())
        .then(data => {
            if (data.done) {
                document.getElementById('status').innerText = data.result;
            } else {
                document.getElementById('status').innerText = '进度: ' + data.progress + '%';
                setTimeout(() => pollStatus(taskID), 2000);
            }
        })
        .catch(err => {
            console.error(err);
            setTimeout(() => pollStatus(taskID), 3000);
        });
}
