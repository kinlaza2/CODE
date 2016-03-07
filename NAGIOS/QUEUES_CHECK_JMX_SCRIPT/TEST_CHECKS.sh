for i in eventsQueue mtsQueue mosQueue  customerCareQueue
do
	for j in MessageCount ConsumerCount MessagesAdded
	do
		./check_HQ $i $j 1
		wait
	done
done
