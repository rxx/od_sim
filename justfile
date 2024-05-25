build: 
    go build -o build/od_sim ./app/

run:
    go run ./app generate_log -sim data/sim.xlsm

test:
    go test ./...
