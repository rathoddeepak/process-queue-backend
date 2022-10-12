//Required files
const express = require('express');
const app = express();
const http = require('http');
const cors = require('cors')
const path = require('path');
const server = http.createServer(app);
const bodyParser = require('body-parser');
const { Server } = require("socket.io");

//Constants
const io = new Server(server, {
    cors: {
        origin: "http://5.181.217.24:3030"
    }
});

const room_prefix = 'dl';
const port = 3030;
const errorResponse = {
	status : 400,
	message : 'Dummy Error!'
}
const commonSuccessReponse = {status : 200};

//Declarations
const roomProgressList = {};

//Functions to keep watch on list
function invalidateGenerateProgress() {
	//Already Working Ignore....
	const rooms = Object.keys(roomProgressList)
	for(let room of rooms){
		const tasks = roomProgressList[room];
		if(tasks?.length > 0){
			for(let task of tasks){
				if(task?.sendingUpdates != true){
					//Ignoring this task as we are already sending updates
					dispatchUpdates(task, room);
				}				
			}
		}else{
			delete roomProgressList[room];
		}
	}
}

function dispatchUpdates(task, room) {
	if(Array.isArray(roomProgressList[room]) == false){
		delete roomProgressList[room];
		invalidateGenerateProgress();
		return
	}
	//Set task as working
	if(task?.sendingUpdates != true){
		setTaskWorking(task, room);
	}	
	//Checking if file is already downloaded;
	if(task.progress >= 100){
		removeTask(task, room);
	}
	let progress = task.progress;
	io.to(room).emit('progress', {...task, progress});
	setTimeout(() => {
		//Check every second for file generating status 
		if(progress >= 100){
			removeTask(task, room);
			//Remove task will call invalidateGenerateProgress
			//as we can check for new task
			return
		}
		progress += 10;
		setTaskProgress(task, room, progress);
		io.to(room).emit('progress', {...task, progress});
		dispatchUpdates(task, room);
	}, 1000);	
}

const removeTask = (task, room) => {
	if(roomProgressList[room]){
		const idx = roomProgressList[room].findIndex(t => t.id == task.id);
		if(idx != -1)roomProgressList[room].splice(idx, 1);
		invalidateGenerateProgress();
	}
}

const setTaskWorking = (task, room) => {
	if(roomProgressList[room]){
		console.log('Working on task id:', task.id);
		const idx = roomProgressList[room].findIndex(t => t.id == task.id);
		if(idx != -1){
			roomProgressList[room][idx].sendingUpdates = true;
		}
	}	
}

const setTaskProgress = (task, room, progress) => {
	if(roomProgressList[room]){
		const idx = roomProgressList[room].findIndex(t => t.id == task.id);
		if(idx != -1){
			roomProgressList[room][idx].progress = progress;
		}
	}	
}


//Server Setup

app.use(cors());
app.use(bodyParser());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/', (req, res) => {
  res.sendFile( __dirname + "/public/" + "index.html" );
});

app.post('/create_new_task', (req, res) => {		
	const {body : {school_id, task}} = req;	
	/*
	 task : {
		name : 'name_of_task',
		id : 'id_of_task', // ID should be unique for identification
	 }
	*/
	if(task){
		task.progress = 0;
		const roomName = `${room_prefix}-${school_id}`;		
		if(roomProgressList[roomName]){
			roomProgressList[roomName].push(task);
		}else{
			roomProgressList[roomName] = [task];
		}
		invalidateGenerateProgress();
		res.send(commonSuccessReponse);		
	}else{
		res.send(errorResponse);
	}
});

app.post('/get_task_list', (req, res) => {
	const {body : {school_id}} = req;
	const roomName = `${room_prefix}-${school_id}`;
	let successResponse = {
		status : 200,
		room_name : roomName,
		list : []
	}
	if(roomProgressList[roomName]){		
		successResponse.list = roomProgressList[roomName]
		res.send(successResponse);
	}else{
		res.send(successResponse);
	}    
});

io.on("connection", (socket) => {
  console.log('New Socket Connection')
  //when client emits event to join room 
  socket.on("join_room", ({room_name}, callback) => {
  	if(room_name != undefined){
  		console.log('Room Joined!')
  		socket.join(room_name);
  	}
  })
});

server.listen(port, () => {
  console.log(`listening on :${port}`);
});