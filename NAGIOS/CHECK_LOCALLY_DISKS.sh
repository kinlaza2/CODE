k=0
problematic_disks=""
counter_of_partitions=0

for i in `df -khl | awk {'print $1'}  | grep "/dev"`
do
is_diskspace_ok=`./check_disk -c 23% -w 21% -t 5 -r $i | grep "DISK OK" | wc -l`
is_diskspace_warn=`./check_disk -c 23% -w 21% -t 5 -r $i | grep "DISK WARNING" | wc -l`
is_diskspace_crit=`./check_disk -c 23% -w 21% -t 5 -r $i | grep "DISK CRITICAL" | wc -l`



if [  $is_diskspace_ok  -eq 1 ]
then
        PARTITION[$counter_of_partitions]=$i
        STATUS[$counter_of_partitions]=$is_diskspace_ok

        if [ $is_diskspace_warn -eq 1 ]
        then
                PARTITION[$counter_of_partitions]=$i
                STATUS[$counter_of_partitions]=$is_diskspace_
        else
                if [ $is_diskspace_crit  -eq 1 ]
                then

                else
                        echo "UNEXPECTED STATUS OF COMMAND"
                        exit 2;
                fi
        fi
        ##echo "Disk Space on Partition $i is NOT OK"
        problematic_disks="$problematic_disks  -    $i"
        k=`expr $k + 1`
fi

counter_of_partitions=`expr $counter_of_partitions + 1`

done

if [ $k -ne 0 ]
then
        echo -e "Disk Space on the following Partitions are below limits:  $problematic_disks \n"
        echo -e "`df -khl`"
        exit 1
else
        echo -e "Disk Space on all partition is OK"
        exit 0;

fi

#echo $problematic_disks
#echo $k


