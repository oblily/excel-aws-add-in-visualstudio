using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Amazon.EC2;
using Amazon.EC2.Model;
using Amazon.ElasticLoadBalancing;
using Amazon.RDS;
using Amazon.RDS.Model;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelAddInAWS
{
    public partial class AWSRibbon
    {
        Excel.Range cell = null;

        Excel.Worksheet activeWorksheet = null;

        int colorIndex = 33;

        int rowCounter = 1;

        List<Instance> instanceList = null;

        AmazonEC2Client ec2Client = null;

        AmazonElasticLoadBalancingClient elbClient = null;

        AmazonRDSClient rdsClient = null;

        private string getStringThroughValue(bool value, int type)
        {
            String str = "";
            if(value)
            {
                if (type == 1)
                {
                    str = "はい"; 
                } else if (type == 2)
                {
                    str = "有効";
                } else if (type == 3)
                {
                    str = "あり";
                }
            } else
            {
                if (type == 1)
                {
                    str = "いいえ";
                }
                else if (type == 2)
                {
                    str = "無効";
                }
                else if (type == 3)
                {
                    str = "なし";
                }
            }
            return str;
        }

        private string getStringThroughTags(List<Amazon.EC2.Model.Tag> tags)
        {
            string name = tags.Count > 0 ? (from tag in tags
                                            where tag.Key == "Name"
                                                   select tag.Value).ToArray()[0] : "-";
            return name;
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void commonAction(string sheetName, bool getEc2Instance)
        {
            #region ■アクティブシート参照の取得とセル変数の宣言
            activeWorksheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

            if (activeWorksheet == null)
            {
                MessageBox.Show("Please open an active work sheet!");
            }
            #endregion

            #region ■シート名の変更
            try
            {
                activeWorksheet.Name = sheetName;
            }
            catch (Exception)
            {
            }
            #endregion

            #region ■シートヘッダーの作成
            rowCounter = 2;

            // シートタイトル
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "システム名： 統合認証基盤更改";

            // 作成日
            rowCounter++;
            rowCounter++;
            int headerStart = rowCounter;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "作成日：";

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "作成者：";

            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "更新日：";

            int headerEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + headerStart + ":" + "C" + headerEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;
            #endregion

            #region ■AWSアクセスクライアントの作成
            if (ec2Client == null)
            {
                ec2Client = new AmazonEC2Client(Amazon.RegionEndpoint.APNortheast1);
            }
            #endregion

            if (getEc2Instance)
            {
                #region ■EC2インスタンス一覧の取得 
                var request = new DescribeInstancesRequest();

                List<Reservation> response = ec2Client.DescribeInstancesAsync(request).Result.Reservations;

                instanceList = new List<Instance>();

                foreach (var reservation in response)
                {
                    instanceList.AddRange(reservation.Instances);
                }
                #endregion
            }
        }

        private void displayEC2InstanceListButton_Click(object sender, RibbonControlEventArgs e)
        {
            #region 共通処理
            commonAction("EC2仕様書", true);
            #endregion

            #region ■EC2インスタンス一覧の描画
            rowCounter = makeInstanceList(instanceList, rowCounter);
            #endregion

            #region ■EC2インスタンスの描画
            rowCounter = makeInstance(instanceList, rowCounter);
            #endregion

            // 列幅を自動調整する
            activeWorksheet.get_Range("B1:M" + rowCounter).Columns.AutoFit();
        }

        private int makeInstanceList(List<Instance> instanceList, int rowCounter)
        {
            // EC2インスタンス一覧エリア
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "■EC2インスタンス一覧";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter + 1;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "インスタンス名";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "インスタンスID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "インスタンスタイプ";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "アベイラビリティーゾーン";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("F" + rowCounter);
            cell.Value2 = "インスタンスの状態";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("G" + rowCounter);
            cell.Value2 = "アラームステータス";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("H" + rowCounter);
            cell.Value2 = "パブリック DNS (IPv4)";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("I" + rowCounter);
            cell.Value2 = "IPv4 パブリック IP";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("J" + rowCounter);
            cell.Value2 = "キー名";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("K" + rowCounter);
            cell.Value2 = "起動時刻";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("L" + rowCounter);
            cell.Value2 = "セキュリティグループ";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("M" + rowCounter);
            cell.Value2 = "所有者";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var instance in instanceList)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = getStringThroughTags(instance.Tags);

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.InstanceId;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = instance.InstanceType.Value;

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = instance.Placement.AvailabilityZone;

                cell = activeWorksheet.get_Range("F" + rowCounter);
                cell.Value2 = instance.State.Name.Value;

                cell = activeWorksheet.get_Range("G" + rowCounter);
                //cell.Value2 = instance.PublicIpAddress;

                cell = activeWorksheet.get_Range("H" + rowCounter);
                cell.Value2 = string.IsNullOrEmpty(instance.PublicDnsName) ? "-" : instance.PublicDnsName;

                cell = activeWorksheet.get_Range("I" + rowCounter);
                cell.Value2 = instance.PublicIpAddress;

                cell = activeWorksheet.get_Range("J" + rowCounter);
                cell.Value2 = instance.KeyName;

                cell = activeWorksheet.get_Range("K" + rowCounter);
                cell.Value2 = instance.LaunchTime;

                cell = activeWorksheet.get_Range("L" + rowCounter);
                List<String> sgList = new List<String>();
                foreach (var reservation in instance.SecurityGroups)
                {
                    sgList.Add(reservation.GroupName);
                }
                cell.Value2 = string.Join(",", sgList);

                cell = activeWorksheet.get_Range("M" + rowCounter);
                cell.Value2 = instance.NetworkInterfaces.Count > 0 ?
                    instance.NetworkInterfaces[0].OwnerId : "-";
                cell.NumberFormatLocal = "@";
            }
            int instanceListEnd = rowCounter - 1;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "M" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeInstance(List<Instance> instanceList, int rowCounter)
        {
            int instanceListStart = rowCounter + 2;

            foreach (var instance in instanceList)
            {
                // EC2インスタンス一覧エリア
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "■EC2インスタンス";

                // タイトル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "項目名";
                cell.Interior.ColorIndex = colorIndex;

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "値";
                cell.Interior.ColorIndex = colorIndex;

                // インスタンスID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "インスタンスID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.InstanceId;

                // インスタンス状態
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "インスタンス状態";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.State.Name.Value;

                // インスタンスタイプ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "インスタンスタイプ";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.InstanceType.ToString();

                // Elastic IP
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "Elastic IP";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "";

                // アベイラビリティーゾーン
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "アベイラビリティーゾーン";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.Placement.AvailabilityZone;

                // セキュリティグループ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "セキュリティグループ";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                List<String> sgList = new List<String>();
                foreach (var reservation in instance.SecurityGroups)
                {
                    sgList.Add(reservation.GroupName);
                }
                cell.Value2 = string.Join(",", sgList);

                // 予定されているイベント
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "予定されているイベント";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "予定されているイベントはありません";

                // AMI ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "AMI ID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.ImageId;

                // プラットフォーム
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "プラットフォーム";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.Platform == null ? "-" : instance.Platform.Value;

                // IAM ロール
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "IAM ロール";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.IamInstanceProfile == null ? "-" : instance.IamInstanceProfile.Id;

                // キーペア名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "キーペア名";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.KeyName;

                // 所有者
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "所有者";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.NetworkInterfaces.Count > 0 ?
                    instance.NetworkInterfaces[0].OwnerId : "-";
                cell.NumberFormatLocal = "@";

                // 起動時刻
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "起動時刻";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.LaunchTime;

                // 終了保護
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "終了保護";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.NetworkInterfaces.Count > 0 ?
                //    instance.NetworkInterfaces[0].Attachment.DeleteOnTermination : false;

                // ライフサイクル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ライフサイクル";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.InstanceLifecycle == null ? "-" : instance.InstanceLifecycle.Value;

                // モニタリング
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "モニタリング";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.Monitoring.State.Value;

                // アラームステータス
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "アラームステータス";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.Monitoring.State.Value;

                // カーネル ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "カーネル ID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = string.IsNullOrEmpty(instance.KernelId) ? "-" : instance.KernelId;

                // RAM ディスク ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "RAM ディスク ID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = string.IsNullOrEmpty(instance.RamdiskId) ? "-" : instance.RamdiskId;

                // 配置グループ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "配置グループ";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = string.IsNullOrEmpty(instance.Placement.GroupName) ? "-" : instance.Placement.GroupName;

                // パーティション番号
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "パーティション番号";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = string.IsNullOrEmpty(instance.Placement.GroupName) ? "-" : instance.Placement.GroupName;

                // 仮想化
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "仮想化";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.VirtualizationType.Value;

                // 予約
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "予約";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.VirtualizationType.Value;

                // AMI 作成インデックス
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "AMI 作成インデックス";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.AmiLaunchIndex;
                cell.NumberFormatLocal = "@";

                // テナンシー
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "テナンシー";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.Placement.Tenancy.Value;

                // ホスト ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ホスト ID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.Placement.Tenancy.Value;

                // アフィニティ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "アフィニティ";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.Placement.Tenancy.Value;

                // 状態遷移の理由
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "状態遷移の理由";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = string.IsNullOrEmpty(instance.StateTransitionReason) ? "-" : instance.StateTransitionReason;

                // 状態遷移の理由メッセージ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "状態遷移の理由メッセージ";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.StateReason == null ? "-" : instance.StateReason.Message;

                // 停止 - 休止動作
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "停止 - 休止動作";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.Placement.Tenancy.Value;

                // vCPU の数
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "vCPU の数";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.Placement.Tenancy.Value;

                // パブリック DNS (IPv4)
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "パブリック DNS (IPv4)";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = string.IsNullOrEmpty(instance.PublicDnsName) ? "-" : instance.PublicDnsName;

                // IPv4 パブリック IP
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "IPv4 パブリック IP";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = string.IsNullOrEmpty(instance.PublicIpAddress) ? "-" : instance.PublicIpAddress;

                // プライベート DNS
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "プライベート DNS";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = string.IsNullOrEmpty(instance.PrivateDnsName) ? "-" : instance.PrivateDnsName;

                // プライベート IP
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "プライベート IP";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = string.IsNullOrEmpty(instance.PrivateIpAddress) ? "-" : instance.PrivateIpAddress;

                // セカンダリプライベート IP
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "セカンダリプライベート IP";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = string.IsNullOrEmpty(instance.PrivateIpAddress) ? "-" : instance.PrivateIpAddress;

                // VPC ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "IPv4 パブリック IP";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.VpcId;

                // サブネット ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "サブネット ID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.SubnetId;

                // ネットワークインターフェイス
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ネットワークインターフェイス";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                List<String> nwList = new List<String>();
                foreach (var reservation in instance.NetworkInterfaces)
                {
                    nwList.Add(reservation.NetworkInterfaceId);  
                }
                cell.Value2 = string.Join(",", nwList);

                // 送信元/送信先チェック
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "送信元/送信先チェック";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.SourceDestCheck;

                // T2/T3 無制限
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "T2/T3 無制限";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.SourceDestCheck;

                // EBS 最適化
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "EBS 最適化";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.EbsOptimized;

                // ルートデバイスタイプ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ルートデバイスタイプ";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.RootDeviceType.Value;

                // ルートデバイス
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ルートデバイス";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.RootDeviceName;

                // ブロックデバイス
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ブロックデバイス";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.BlockDeviceMappings[0].DeviceName;

                // Elastic Graphics ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "Elastic Graphics ID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.Architecture.Value;

                // Elastic Inference アクセラレーター ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "Elastic Inference アクセラレーター ID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.Architecture.Value;

                // キャパシティーの予約
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "キャパシティーの予約";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.Architecture.Value;

                // キャパシティー予約設定
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "キャパシティー予約設定";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.Architecture.Value;

                // Outpost ARN
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "Outpost ARN";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = instance.Architecture.Value;

                // Architecture
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "Architecture";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.Architecture.Value;
            }

            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "C" + instanceListEnd).Borders.LineStyle = true;

            return rowCounter;
        }

        private void displaySecurityGroupButton_Click(object sender, RibbonControlEventArgs e)
        {
            #region 共通処理
            commonAction("SecurityGroup仕様書", true);
            #endregion

            #region ■Security Group一覧の取得
            DescribeSecurityGroupsRequest sgRequest = new DescribeSecurityGroupsRequest();

            List<SecurityGroup> securityGroups = ec2Client.DescribeSecurityGroupsAsync(sgRequest).Result.SecurityGroups;
            #endregion

            #region ■Security Group一覧の描画
            rowCounter = makeSecurityGroupList(securityGroups, rowCounter);
            #endregion

            #region ■Security Groupの描画
            rowCounter = makeSecurityGroup(securityGroups, rowCounter);
            #endregion

            // 列幅を自動調整する
            activeWorksheet.get_Range("B1:F" + rowCounter).Columns.AutoFit();
        }

        private int makeSecurityGroupList(List<SecurityGroup> sgList, int rowCounter)
        {
            // セキュリティグループ一覧エリア
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "■セキュリティグループ一覧";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "Name";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "グループID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "グループ名";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "VPC ID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("F" + rowCounter);
            cell.Value2 = "説明";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var sg in sgList)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = getStringThroughTags(sg.Tags);

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = sg.GroupId;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = sg.GroupName;

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = sg.VpcId;

                cell = activeWorksheet.get_Range("F" + rowCounter);
                cell.Value2 = sg.Description;
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "F" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeSecurityGroup(List<SecurityGroup> sgList, int rowCounter)
        {
            foreach (var sg in sgList)
            {
                int instanceListStart = rowCounter + 2;

                // Security Group詳細タイトル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "■セキュリティグループ";

                // Security Group詳細ヘッダー
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "項目名";
                cell.Interior.ColorIndex = colorIndex;

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "値";
                cell.Interior.ColorIndex = colorIndex;

                // グループ名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "グループ名";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = sg.GroupName;

                // グループ ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "グループ ID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = sg.GroupId;

                // VPC ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "VPC ID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = sg.VpcId;

                // 説明
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "説明";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = sg.Description;

                // インバウンドのルールカウント
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "インバウンドのルールカウント";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = sg.IpPermissions.Count;
                cell.NumberFormatLocal = "@";

                // アウトバウンドのルールカウント
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "アウトバウンドのルールカウント";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = sg.IpPermissionsEgress.Count;
                cell.NumberFormatLocal = "@";

                // 所有者
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "所有者";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = sg.OwnerId;
                cell.NumberFormatLocal = "@";

                int instanceListEnd = rowCounter;

                // インスタンス一覧に罫線を引く
                activeWorksheet.get_Range("B" + instanceListStart + ":" + "C" + instanceListEnd).Borders.LineStyle = true;

                // 1行開ける
                rowCounter++;

                #region ■インバウンドのルールの描画
                rowCounter = makeInboundForSecurityGroup(sg.IpPermissions, rowCounter);
                #endregion

                #region ■アウトバウンドのルールの描画
                rowCounter = makeOutboundForSecurityGroup(sg.IpPermissionsEgress, rowCounter);
                #endregion
            }
            return rowCounter;
        }

        private int makeInboundForSecurityGroup(List<IpPermission> ipPermission, int rowCounter)
        {
            // インバウンドのルールエリア
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "▼インバウンドのルール";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "タイプ";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "プロトコル";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "ポート範囲";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "ソース";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("F" + rowCounter);
            cell.Value2 = "説明";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var rule in ipPermission)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = rule.FromPort == 0 ?
                            (rule.IpProtocol.Equals("-1") ? "すべてのトラフィック" : "すべての" + rule.IpProtocol) :
                            rule.FromPort.ToString();

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rule.IpProtocol.Equals("-1") ?
                         "すべて" : rule.IpProtocol;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = rule.ToPort == 0 ? "すべて" : rule.ToPort.ToString();

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = rule.IpRanges.Count > 0 ?
                           rule.IpRanges[0] :
                           (rule.UserIdGroupPairs.Count > 0 ?
                           rule.UserIdGroupPairs[0].GroupId : "");

                cell = activeWorksheet.get_Range("F" + rowCounter);
                cell.Value2 = "";
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "F" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeOutboundForSecurityGroup(List<IpPermission> ipPermissionEgress, int rowCounter)
        {
            // アウトバウンドのルールエリア
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "▼アウトバウンドのルール";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "タイプ";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "プロトコル";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "ポート範囲";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "送信先";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("F" + rowCounter);
            cell.Value2 = "説明";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var rule in ipPermissionEgress)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = rule.FromPort == 0 ?
                            (rule.IpProtocol.Equals("-1") ? "すべてのトラフィック" : "すべての" + rule.IpProtocol) :
                            rule.FromPort.ToString();

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rule.IpProtocol.Equals("-1") ?
                         "すべて" : rule.IpProtocol;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = rule.ToPort == 0 ? "すべて" : rule.ToPort.ToString();

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = rule.IpRanges.Count > 0 ?
                           rule.IpRanges[0] :
                           (rule.UserIdGroupPairs.Count > 0 ?
                           rule.UserIdGroupPairs[0].GroupId : "");

                cell = activeWorksheet.get_Range("F" + rowCounter);
                cell.Value2 = "";
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "F" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private void displayElbButton_Click(object sender, RibbonControlEventArgs e)
        {
            #region 共通処理
            commonAction("ELB仕様書", false);
            #endregion

            #region ■elbインスタンス         
            elbClient = new AmazonElasticLoadBalancingClient(Amazon.RegionEndpoint.APNortheast1);

            #region ■elbインスタンス一覧の取得 
            var request = new Amazon.ElasticLoadBalancing.Model.DescribeLoadBalancersRequest();

            List<Amazon.ElasticLoadBalancing.Model.LoadBalancerDescription> response = elbClient.DescribeLoadBalancersAsync(request).Result.LoadBalancerDescriptions;

            #endregion

            #region ■elbインスタンス一覧の描画
            rowCounter = makeElbList(response, rowCounter);
            #endregion

            #region ■elbインスタンスの描画
            rowCounter = makeElb(response, rowCounter);
            #endregion

            // 列幅を自動調整する
            activeWorksheet.get_Range("B1:H" + rowCounter).Columns.AutoFit();
            #endregion
        }

        private int makeElbList(List<Amazon.ElasticLoadBalancing.Model.LoadBalancerDescription> elbList, int rowCounter)
        {
            // ELB一覧エリア
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "■ロードバランサー一覧";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "ロードバランサー名";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "DNS名";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "状態";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "VPC ID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("F" + rowCounter);
            cell.Value2 = "アベイラビリティーゾーン";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("G" + rowCounter);
            cell.Value2 = "種類";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("H" + rowCounter);
            cell.Value2 = "作成日時";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var elb in elbList)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = elb.LoadBalancerName;

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = elb.DNSName;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                //cell.Value2 = elb.CreatedTime;

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = elb.VPCId;

                cell = activeWorksheet.get_Range("F" + rowCounter);
                cell.Value2 = String.Join(",", elb.AvailabilityZones);

                cell = activeWorksheet.get_Range("G" + rowCounter);
                //cell.Value2 = elb.CreatedTime;

                cell = activeWorksheet.get_Range("H" + rowCounter);
                cell.Value2 = elb.CreatedTime;
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "H" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeElb(List<Amazon.ElasticLoadBalancing.Model.LoadBalancerDescription> elbList, int rowCounter)
        {
            foreach (var elb in elbList)
            {
                int instanceListStart = rowCounter + 2;

                // ロードバランサエリア
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "■ロードバランサ";

                // タイトル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "項目名";
                cell.Interior.ColorIndex = colorIndex;

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "値";
                cell.Interior.ColorIndex = colorIndex;

                // ロードバランサー名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ロードバランサー名";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = elb.LoadBalancerName;

                // ARN
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ARN";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = elb.DNSName;

                // DNS名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "DNS名";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = elb.DNSName;

                // 状態
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "状態";
                /*cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = elb.DNSName;*/

                // 種類
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "種類";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = elb.Scheme;

                // スキーム
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "スキーム";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = elb.Scheme;

                // IP アドレスタイプ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "IP アドレスタイプ";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "ipv4";

                // VPC
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "VPC";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = elb.VPCId;

                // アベイラビリティーゾーン
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "アベイラビリティーゾーン";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = String.Join("\n", elb.AvailabilityZones);

                // ホストゾーン
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ホストゾーン";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = elb.CanonicalHostedZoneNameID;

                // 作成時刻
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "作成時刻";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = elb.CreatedTime;

                // セキュリティグループ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "セキュリティグループ";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = String.Join("\n", elb.SecurityGroups);

                int instanceListEnd = rowCounter;

                // インスタンス一覧に罫線を引く
                activeWorksheet.get_Range("B" + instanceListStart + ":" + "C" + instanceListEnd).Borders.LineStyle = true;
                rowCounter++;

                #region ■リスナーの描画
                rowCounter = makeListenerForElb(elb, rowCounter);
                #endregion

                #region ■EC2インスタンスの描画
                rowCounter = makeInstanceForElb(elb, rowCounter);
                #endregion

                #region ■アベイラビリティーゾーンの描画
                rowCounter = makeAzForElb(elb, rowCounter);
                #endregion

                #region ■ヘルスチェックの描画
                rowCounter = makeHealthCheckForElb(elb, rowCounter);
                #endregion
            }
            return rowCounter;
        }

        private int makeListenerForElb(Amazon.ElasticLoadBalancing.Model.LoadBalancerDescription elb, int rowCounter)
        {
            // インスタンスエリア
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "▼リスナー";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "リスナー ID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "セキュリティポリシー";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "SSL 証明書";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "ルール";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var listener in elb.ListenerDescriptions)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = listener.Listener.Protocol + ":" + listener.Listener.LoadBalancerPort.ToString();

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = listener.PolicyNames.Count > 0 ?
                    String.Join(",", listener.PolicyNames) : "該当なし";

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = string.IsNullOrEmpty(listener.Listener.SSLCertificateId) ? "該当なし" : listener.Listener.SSLCertificateId;

                cell = activeWorksheet.get_Range("E" + rowCounter);
                //cell.Value2 = vpc.CidrBlock;
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "D" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeInstanceForElb(Amazon.ElasticLoadBalancing.Model.LoadBalancerDescription elb, int rowCounter)
        {
            // インスタンスエリア
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "▼インスタンス";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "インスタンス名";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "インスタンスID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "アベイラビリティーゾーン";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var instance in elb.Instances)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = instance.InstanceId;

                /*cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpc.VpcId;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = vpc.State.Value;

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = vpc.CidrBlock;

                cell = activeWorksheet.get_Range("F" + rowCounter);
                cell.Value2 = vpc.DhcpOptionsId;

                cell = activeWorksheet.get_Range("G" + rowCounter);
                cell.Value2 = vpc.IsDefault;*/
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "D" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeAzForElb(Amazon.ElasticLoadBalancing.Model.LoadBalancerDescription elb, int rowCounter)
        {
            // アベイラビリティーゾーンエリア
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "▼アベイラビリティーゾーン";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "アベイラビリティーゾーン";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "サブネットID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "サブネットCIDR";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "ターゲット数";
            cell.Interior.ColorIndex = colorIndex;

            foreach (string subnet in elb.Subnets)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = subnet;

                /*cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpc.VpcId;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = vpc.State.Value;

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = vpc.CidrBlock;

                cell = activeWorksheet.get_Range("F" + rowCounter);
                cell.Value2 = vpc.DhcpOptionsId;

                cell = activeWorksheet.get_Range("G" + rowCounter);
                cell.Value2 = vpc.IsDefault;*/
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "E" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeHealthCheckForElb(Amazon.ElasticLoadBalancing.Model.LoadBalancerDescription elb, int rowCounter)
        {
            // ヘルスチェックエリア
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "▼ヘルスチェック方式";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "項目名";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "値";
            cell.Interior.ColorIndex = colorIndex;

            // ターゲット
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "ターゲット";
            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = elb.HealthCheck.Target;

            // 正常のしきい値
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "正常のしきい値";
            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = elb.HealthCheck.HealthyThreshold;

            // 非正常のしきい値
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "非正常のしきい値";
            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = elb.HealthCheck.UnhealthyThreshold;

            // タイムアウト
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "タイムアウト";
            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = elb.HealthCheck.Timeout;

            // 間隔
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "間隔";
            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = elb.HealthCheck.Interval;

            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "C" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private void displayVPCListButton_Click(object sender, RibbonControlEventArgs e)
        {
            #region 共通処理
            commonAction("VPC仕様書", false);
            #endregion

            #region ■VPC一覧の取得 
            var request = new DescribeVpcsRequest();

            List<Vpc> response = ec2Client.DescribeVpcsAsync(request).Result.Vpcs;
            #endregion

            #region ■VPC一覧の描画
            rowCounter = makeVpcList(response, rowCounter);
            #endregion

            #region ■VPCの描画
            rowCounter = makeVpc(response, rowCounter);
            #endregion

            // 列幅を自動調整する
            activeWorksheet.get_Range("B1:H" + rowCounter).Columns.AutoFit();     
        }

        private int makeVpcList(List<Vpc> vpcList, int rowCounter)
        {
            // VPC一覧エリア
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "■VPC一覧";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "VPC名";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "VPC ID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "状態";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "IPv4 CIDR";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("F" + rowCounter);
            cell.Value2 = "デフォルト VPC";
            cell.Interior.ColorIndex = colorIndex;
            foreach (var vpc in vpcList)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = getStringThroughTags(vpc.Tags);

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpc.VpcId;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = vpc.State.Value;

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = vpc.CidrBlock;

                cell = activeWorksheet.get_Range("F" + rowCounter);
                cell.Value2 = getStringThroughValue(vpc.IsDefault, 1);
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "F" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeVpc(List<Vpc> vpcList, int rowCounter)
        {
            foreach (var vpc in vpcList)
            {
                int instanceListStart = rowCounter + 2;

                // VPCエリア
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "■VPC";

                // タイトル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "項目名";
                cell.Interior.ColorIndex = colorIndex;

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "値";
                cell.Interior.ColorIndex = colorIndex;

                // VPC名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "VPC名";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = getStringThroughTags(vpc.Tags);

                // VPC ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "VPC ID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpc.VpcId;

                // 状態
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "状態";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpc.State.Value;

                // IPv4 CIDR
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "IPv4 CIDR";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpc.CidrBlock;

                // ネットワーク ACL
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ネットワーク ACL";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                #region ■ネットワーク ACL取得
                var defaultRtFilter = new Amazon.EC2.Model.Filter { Name = "default" };
                defaultRtFilter.Values.Add("true");
                var aclFilter = new Amazon.EC2.Model.Filter { Name = "vpc-id" };
                aclFilter.Values.Add(vpc.VpcId);
                var request = new DescribeNetworkAclsRequest();
                request.Filters.Add(defaultRtFilter);
                request.Filters.Add(aclFilter);
                var networkAcls = ec2Client.DescribeNetworkAcls(request);
                if (networkAcls.NetworkAcls.Any())
                {
                    cell.Value2 = networkAcls.NetworkAcls[0].NetworkAclId;
                }
                #endregion

                // DHCP オプションセット
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "DHCP オプションセット";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpc.DhcpOptionsId;

                // ルートテーブル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ルートテーブル";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                #region ■ルートテーブル取得
                var mainRtFilter = new Amazon.EC2.Model.Filter { Name = "association.main" };
                mainRtFilter.Values.Add("true");
                var rtRequest = new DescribeRouteTablesRequest();
                rtRequest.Filters.Add(aclFilter);
                rtRequest.Filters.Add(mainRtFilter);
                var routeTables = ec2Client.DescribeRouteTables(rtRequest);
                if (routeTables.RouteTables.Any())
                {
                    cell.Value2 = routeTables.RouteTables[0].RouteTableId;
                }
                #endregion

                // テナンシー
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "テナンシー";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpc.InstanceTenancy.Value;

                // デフォルト VPC
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "デフォルト VPC";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = getStringThroughValue(vpc.IsDefault,1);

                // ClassicLink
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ClassicLink";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = getStringThroughValue(false, 2);

                #region ■DNS属性取得 
                var attrDnsSupportRequest = new DescribeVpcAttributeRequest
                {
                    Attribute = "enableDnsSupport",
                    VpcId = vpc.VpcId
                };
                var attrDnsHostnamesRequest = new DescribeVpcAttributeRequest
                {
                    Attribute = "enableDnsHostnames",
                    VpcId = vpc.VpcId
                };
                bool enableDnsSupport = ec2Client.DescribeVpcAttribute(attrDnsSupportRequest).EnableDnsSupport;
                bool enableDnsHostnames = ec2Client.DescribeVpcAttribute(attrDnsHostnamesRequest).EnableDnsHostnames;
                // DNS 解決
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "DNS 解決";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = getStringThroughValue(enableDnsSupport, 2);

                // DNS ホスト名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "DNS ホスト名";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = getStringThroughValue(enableDnsHostnames, 2);
                #endregion

                // ClassicLink DNS サポート
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ClassicLink DNS サポート";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = getStringThroughValue(false, 2);

                // 所有者
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "所有者";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = request.Filters[1].Values[0];

                int instanceListEnd = rowCounter;

                // インスタンス一覧に罫線を引く
                activeWorksheet.get_Range("B" + instanceListStart + ":" + "C" + instanceListEnd).Borders.LineStyle = true;
                rowCounter++;
            }
            return rowCounter;
        }

        private void displayRdsButton_Click(object sender, RibbonControlEventArgs e)
        {
            #region 共通処理
            commonAction("RDS仕様書", false);
            #endregion

            #region ■rdsインスタンス
            if (rdsClient == null)
            {
                rdsClient = new AmazonRDSClient(Amazon.RegionEndpoint.APNortheast1);

                #region ■rdsインスタンス一覧の取得 
                var request = new DescribeDBInstancesRequest();

                List<DBInstance> response = rdsClient.DescribeDBInstancesAsync(request).Result.DBInstances;
                #endregion

                #region ■rdsインスタンス一覧の描画
                rowCounter = makeRdsList(response, rowCounter);
                #endregion

                #region ■rdsインスタンスの描画
                rowCounter = makeRds(response, rowCounter);
                #endregion

                // 列幅を自動調整する
                activeWorksheet.get_Range("B1:J" + rowCounter).Columns.AutoFit();
            }
            #endregion
        }

        private int makeRdsList(List<DBInstance> rdsList, int rowCounter)
        {
            // DB一覧エリア
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "■DBインスタンス一覧";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "DB 識別子";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "ロール";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "エンジン";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "リージョンと AZ";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("F" + rowCounter);
            cell.Value2 = "サイズ";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("G" + rowCounter);
            cell.Value2 = "ストレージ";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("H" + rowCounter);
            cell.Value2 = "VPC セキュリティグループ";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("I" + rowCounter);
            cell.Value2 = "VPC ID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("J" + rowCounter);
            cell.Value2 = "マルチ AZ";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var rds in rdsList)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = rds.DBInstanceIdentifier;

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "インスタンス";

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = rds.Engine;

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = rds.AvailabilityZone;

                cell = activeWorksheet.get_Range("F" + rowCounter);
                cell.Value2 = rds.DBInstanceClass;

                cell = activeWorksheet.get_Range("G" + rowCounter);
                cell.Value2 = rds.AllocatedStorage;

                cell = activeWorksheet.get_Range("H" + rowCounter);
                List<String> dbSg = new List<String>();
                foreach (var sg in rds.DBSecurityGroups)
                {
                    dbSg.Add(sg.DBSecurityGroupName);
                }
                cell.Value2 = string.Join(",", dbSg);

                cell = activeWorksheet.get_Range("I" + rowCounter);
                cell.Value2 = rds.DBSubnetGroup.VpcId;

                cell = activeWorksheet.get_Range("J" + rowCounter);
                cell.Value2 = getStringThroughValue(rds.MultiAZ, 3);
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "J" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeRds(List<DBInstance> rdsList, int rowCounter)
        {
            foreach (var rds in rdsList)
            {
                int instanceListStart = rowCounter + 2;

                // DBインスタンスエリア
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "■ DBインスタンス";

                // タイトル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "項目名";
                cell.Interior.ColorIndex = colorIndex;

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "値";
                cell.Interior.ColorIndex = colorIndex;

                // DBインスタンス名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "DBインスタンス名";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.DBInstanceIdentifier;

                // エンドポイント
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "エンドポイント";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.Endpoint.Address;

                // ポート
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ポート";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.Endpoint.Port;

                // エンジン
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "エンジン";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.Engine;

                // エンジンバージョン
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "エンジンバージョン";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = rds.Engine;

                // DB 名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "DB 名";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.DBName;

                // ライセンスモデル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ライセンスモデル";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.LicenseModel;

                // オプショングループ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "オプショングループ";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                List<String> optinGroups = new List<String>();
                foreach (var opg in rds.OptionGroupMemberships)
                {
                    optinGroups.Add(opg.OptionGroupName);
                }
                cell.Value2 = string.Join(",", optinGroups);

                // ARN
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ARN";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = rds.InstanceCreateTime;

                // リソース ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "リソース ID";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = rds.InstanceCreateTime;

                // 作成時刻
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "VPC";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.InstanceCreateTime;

                // パラメータグループ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "パラメータグループ";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                List<String> paramGroups = new List<String>();
                foreach (var parg in rds.DBParameterGroups)
                {
                    paramGroups.Add(parg.DBParameterGroupName);
                }
                cell.Value2 = string.Join(",", paramGroups);

                // 削除保護
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "削除保護";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "有効";

                // インスタンスクラス
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "インスタンスクラス";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.DBInstanceClass;

                // マスターユーザー名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "マスターユーザー名";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.MasterUsername;

                // IAM db 認証
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "IAM db 認証";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "有効でない";

                // マルチ AZ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "マルチ AZ";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = getStringThroughValue(rds.MultiAZ, 3);

                // セカンダリゾーン
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "セカンダリゾーン";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.SecondaryAvailabilityZone;

                // 暗号化
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "暗号化";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = rds.StorageType;
               
                // ストレージタイプ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ストレージタイプ";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = /*rds.StorageType*/"汎用 (SSD)";

                // IOPS
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "IOPS";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.Iops;

                // ストレージ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ストレージ";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.AllocatedStorage;

                // ストレージの自動スケーリング
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ストレージの自動スケーリング";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = rds.AllocatedStorage;

                // Performance Insights が有効
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "Performance Insights が有効";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = rds.AllocatedStorage;

                // 最大ストレージしきい値
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "最大ストレージしきい値";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = rds.AllocatedStorage;

                // アベイラビリティーゾーン
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "アベイラビリティーゾーン";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.AvailabilityZone;

                // VPC ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "VPC ID";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.DBSubnetGroup.VpcId;

                // サブネットグループ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "サブネットグループ";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.DBSubnetGroup.DBSubnetGroupName;

                // サブネット
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "サブネット";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                List<String> subnets = new List<String>();
                foreach (var subnet in rds.DBSubnetGroup.Subnets)
                {
                    subnets.Add(subnet.SubnetIdentifier);
                }
                cell.Value2 = string.Join("\n", subnets);

                // VPC セキュリティグループ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "VPC セキュリティグループ";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                List<String> vpcSg = new List<String>();
                foreach (var sg in rds.VpcSecurityGroups)
                {
                    vpcSg.Add(sg.VpcSecurityGroupId);
                }
                cell.Value2 = string.Join("\n", vpcSg);

                // パブリックアクセシビリティ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "パブリックアクセシビリティ";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.PubliclyAccessible;

                // 認証機関
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "認証機関";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = rds.PubliclyAccessible;

                // 証明機関の日付
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "証明機関の日付";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = rds.PubliclyAccessible;

                // マイナーバージョン自動アップグレード
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "マイナーバージョン自動アップグレード";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.AutoMinorVersionUpgrade;

                // メンテナンスウィンドウ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "メンテナンスウィンドウ";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.PreferredMaintenanceWindow;

                // 自動バックアップ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "自動バックアップ";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = rds.PreferredBackupWindow;

                // スナップショットにタグをコピー
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "スナップショットにタグをコピー";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = rds.PreferredBackupWindow;

                // バックアップウィンドウ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "バックアップウィンドウ";
                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = rds.PreferredBackupWindow;

                int instanceListEnd = rowCounter;

                // インスタンス一覧に罫線を引く
                activeWorksheet.get_Range("B" + instanceListStart + ":" + "C" + instanceListEnd).Borders.LineStyle = true;
                rowCounter++;
            }
            return rowCounter;
        }

        private void displaySubnetButton_Click(object sender, RibbonControlEventArgs e)
        {
            #region 共通処理
            commonAction("Subnet仕様書", false);
            #endregion

            #region ■サブネット一覧の取得 
            var request = new DescribeSubnetsRequest();

            List<Amazon.EC2.Model.Subnet> response = ec2Client.DescribeSubnetsAsync(request).Result.Subnets;
            #endregion

            #region ■サブネット一覧の描画
            rowCounter = makeSubnetList(response, rowCounter);
            #endregion

            #region ■サブネットの描画
            rowCounter = makeSubnet(response, rowCounter);
            #endregion

            // 列幅を自動調整する
            activeWorksheet.get_Range("B1:H" + rowCounter).Columns.AutoFit();
        }

        private int makeSubnetList(List<Amazon.EC2.Model.Subnet> subnetList, int rowCounter)
        {
            // サブネット一覧エリア
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "■サブネット一覧";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "サブネット名";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "サブネットID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "状態";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "VPC ID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("F" + rowCounter);
            cell.Value2 = "IPv4 CIDR";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("G" + rowCounter);
            cell.Value2 = "アベイラビリティーゾーン";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var subnet in subnetList)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = getStringThroughTags(subnet.Tags);

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = subnet.SubnetId;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = subnet.State.Value;

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = subnet.VpcId;

                cell = activeWorksheet.get_Range("F" + rowCounter);
                cell.Value2 = subnet.CidrBlock;

                cell = activeWorksheet.get_Range("G" + rowCounter);
                cell.Value2 = subnet.AvailabilityZone;
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "G" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeSubnet(List<Amazon.EC2.Model.Subnet> subnetList, int rowCounter)
        {
            foreach (var subnet in subnetList)
            {
                int instanceListStart = rowCounter + 2;

                // サブネットエリア
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "■サブネット";

                // タイトル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "項目名";
                cell.Interior.ColorIndex = colorIndex;

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "値";
                cell.Interior.ColorIndex = colorIndex;

                // サブネット名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "サブネット名";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = getStringThroughTags(subnet.Tags); ;

                // サブネットID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "サブネットID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = subnet.SubnetId;

                // 状態
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "状態";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = subnet.State.Value;

                // VPC
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "VPC";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = subnet.VpcId;

                // IPv4 CIDR
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "IPv4 CIDR";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = subnet.CidrBlock;

                // 利用可能な IPv4 アドレス
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "利用可能な IPv4 アドレス";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = subnet.AvailableIpAddressCount;

                // アベイラビリティーゾーン
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "アベイラビリティーゾーン";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = subnet.AvailabilityZone;

                // ルートテーブル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ルートテーブル";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                #region ■ルートテーブル取得
                var rtFilter = new Amazon.EC2.Model.Filter { Name = "association.subnet-id" };
                rtFilter.Values.Add(subnet.SubnetId);
                var rtRequest = new DescribeRouteTablesRequest();
                rtRequest.Filters.Add(rtFilter);
                var routeTables = ec2Client.DescribeRouteTables(rtRequest);
                if (routeTables.RouteTables.Any())
                {
                    cell.Value2 = routeTables.RouteTables[0].RouteTableId;
                }
                #endregion

                // ネットワーク ACL
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ネットワーク ACL";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                #region ■ネットワーク ACL取得
                var aclRequest = new DescribeNetworkAclsRequest();
                aclRequest.Filters.Add(rtFilter);
                var networkAcls = ec2Client.DescribeNetworkAcls(aclRequest);
                if (networkAcls.NetworkAcls.Any())
                {
                    cell.Value2 = networkAcls.NetworkAcls[0].NetworkAclId;
                }
                #endregion

                // デフォルトのサブネット
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "デフォルトのサブネット";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = subnet.DefaultForAz;

                // パブリック IPv4 アドレスの自動
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "パブリック IPv4 アドレスの自動";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = subnet.MapPublicIpOnLaunch;

                // 所有者
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "所有者";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = request.Filters[1].Values[0];

                int instanceListEnd = rowCounter;

                // インスタンス一覧に罫線を引く
                activeWorksheet.get_Range("B" + instanceListStart + ":" + "C" + instanceListEnd).Borders.LineStyle = true;
                rowCounter++;
            }

            return rowCounter;
        }

        private void displayRouteTableButton_Click(object sender, RibbonControlEventArgs e)
        {
            #region 共通処理
            commonAction("RouteTable仕様書", false);
            #endregion

            #region ■ルートテーブル一覧の取得 
            var request = new DescribeRouteTablesRequest();

            List<RouteTable> response = ec2Client.DescribeRouteTablesAsync(request).Result.RouteTables;
            #endregion

            #region ■ルートテーブル一覧の描画
            rowCounter = makeRouteTableList(response, rowCounter);
            #endregion

            #region ■ルートテーブルの描画
            rowCounter = makeRouteTable(response, rowCounter);
            #endregion

            // 列幅を自動調整する
            activeWorksheet.get_Range("B1:H" + rowCounter).Columns.AutoFit();
        }

        private int makeRouteTableList(List<RouteTable> routeTableList, int rowCounter)
        {
            // ルートテーブル一覧エリア
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "■ルートテーブル一覧";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "ルートテーブル名";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "ルートテーブルID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "明示的に関連付けられた";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "VPC ID";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var routeTable in routeTableList)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = getStringThroughTags(routeTable.Tags);

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = routeTable.RouteTableId;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = routeTable.Associations.Count > 0 ?
                            routeTable.Associations.Count.ToString() + "個のサブネット" : "-";

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = routeTable.VpcId;
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "E" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeRouteTable(List<RouteTable> routeTableList, int rowCounter)
        {
            foreach (var routeTable in routeTableList)
            {
                int instanceListStart = rowCounter + 2;

                // ルートテーブルエリア
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "■ルートテーブル";

                // タイトル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "項目名";
                cell.Interior.ColorIndex = colorIndex;

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "値";
                cell.Interior.ColorIndex = colorIndex;

                // ルートテーブル名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ルートテーブル名";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = getStringThroughTags(routeTable.Tags);

                // ルートテーブルID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ルートテーブルID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = routeTable.RouteTableId;

                // 明示的に関連付けられた
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "明示的に関連付けられた";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = routeTable.Associations.Count > 0 ?
                            routeTable.Associations.Count.ToString() + "個のサブネット" : "-";

                // メイン
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "メイン";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                #region ■ルートテーブル取得
                var rtFilter = new Amazon.EC2.Model.Filter { Name = "route-table-id" };
                rtFilter.Values.Add(routeTable.RouteTableId);
                var rtRequest = new DescribeRouteTablesRequest();
                rtRequest.Filters.Add(rtFilter);
                var routeTables = ec2Client.DescribeRouteTables(rtRequest);
                if (routeTables.RouteTables.Any())
                {
                    cell.Value2 = getStringThroughValue(routeTables.RouteTables[0].Associations[0].Main, 1);
                }
                #endregion

                // VPC
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "VPC";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = routeTable.VpcId;

                // 所有者
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "所有者";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = routeTable.VpcId;

                int instanceListEnd = rowCounter;

                // インスタンス一覧に罫線を引く
                activeWorksheet.get_Range("B" + instanceListStart + ":" + "C" + instanceListEnd).Borders.LineStyle = true;
                rowCounter++;

                #region ■ルートの描画
                rowCounter = makeRouteForRouteTable(routeTable.Routes, rowCounter);
                #endregion

                #region ■サブネットの描画
                rowCounter = makeSubnetForRouteTable(routeTable.Associations, rowCounter);
                #endregion
            }

            return rowCounter;
        }

        private int makeRouteForRouteTable(List<Route> route, int rowCounter)
        {
            // ルートの関連付けエリア
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "▼ルート";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "送信先";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "ターゲット";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "ステータス";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "伝播済み";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var instance in route)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = instance.DestinationCidrBlock;

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = instance.GatewayId;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = instance.State.Value;

                //cell = activeWorksheet.get_Range("E" + rowCounter);
                //cell.Value2 = vpc.CidrBlock;
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "E" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeSubnetForRouteTable(List<RouteTableAssociation> rtAssociation, int rowCounter)
        {
            // サブネットの関連付けエリア
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "▼サブネットの関連付け";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "サブネットID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "IPv4 CIDR";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var instance in rtAssociation)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = instance.SubnetId;

                /*cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpc.VpcId;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = vpc.State.Value;

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = vpc.CidrBlock;

                cell = activeWorksheet.get_Range("F" + rowCounter);
                cell.Value2 = vpc.DhcpOptionsId;

                cell = activeWorksheet.get_Range("G" + rowCounter);
                cell.Value2 = vpc.IsDefault;*/
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "C" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private void displayIgwButton_Click(object sender, RibbonControlEventArgs e)
        {
            #region 共通処理
            commonAction("IGW仕様書", false);
            #endregion

            #region ■インターネットゲートウェイ一覧の取得 
            var request = new DescribeInternetGatewaysRequest();

            List<InternetGateway> response = ec2Client.DescribeInternetGatewaysAsync(request).Result.InternetGateways;
            #endregion

            #region ■インターネットゲートウェイ一覧の描画
            rowCounter = makeIgwList(response, rowCounter);
            #endregion

            #region ■インターネットゲートウェイの描画
            rowCounter = makeIgw(response, rowCounter);
            #endregion

            // 列幅を自動調整する
            activeWorksheet.get_Range("B1:H" + rowCounter).Columns.AutoFit();
        }

        private int makeIgwList(List<InternetGateway> igwList, int rowCounter)
        {
            // インターネットゲートウェイ一覧エリア
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "■インターネットゲートウェイ一覧";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "インターネットゲートウェイ名";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "インターネットゲートウェイID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "状態";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "VPC ID";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var igw in igwList)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = getStringThroughTags(igw.Tags);

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = igw.InternetGatewayId;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = igw.Attachments[0].State.Value;

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = igw.Attachments[0].VpcId;
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "E" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeIgw(List<InternetGateway> igwList, int rowCounter)
        {
            foreach (var igw in igwList)
            {
                int instanceListStart = rowCounter + 2;

                // インターネットゲートウェイエリア
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "■インターネットゲートウェイ";

                // タイトル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "項目名";
                cell.Interior.ColorIndex = colorIndex;

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "値";
                cell.Interior.ColorIndex = colorIndex;

                // インターネットゲートウェイ名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "インターネットゲートウェイ名";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = getStringThroughTags(igw.Tags);

                // インターネットゲートウェイID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "インターネットゲートウェイID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = igw.InternetGatewayId;

                // 状態
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "状態";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = igw.Attachments[0].State.Value;

                // アタッチ済み VPC ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "アタッチ済み VPC ID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = igw.Attachments[0].VpcId;

                int instanceListEnd = rowCounter;

                // インスタンス一覧に罫線を引く
                activeWorksheet.get_Range("B" + instanceListStart + ":" + "C" + instanceListEnd).Borders.LineStyle = true;
                rowCounter++;
            }

            return rowCounter;
        }

        private void displayVpnButton_Click(object sender, RibbonControlEventArgs e)
        {
            #region 共通処理
            commonAction("VPN仕様書", false);
            #endregion

            #region ■VPN一覧の取得 
            var request = new DescribeVpnConnectionsRequest();

            List<VpnConnection> response = ec2Client.DescribeVpnConnectionsAsync(request).Result.VpnConnections;
            #endregion

            #region ■VPN一覧の描画
            rowCounter = makeVpnList(response, rowCounter);
            #endregion

            #region ■VPNの描画
            rowCounter = makeVpn(response, rowCounter);
            #endregion

            // 列幅を自動調整する
            activeWorksheet.get_Range("B1:H" + rowCounter).Columns.AutoFit();
        }

        private int makeVpnList(List<VpnConnection> vpnList, int rowCounter)
        {
            // VPN一覧エリア
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "■VPN一覧";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "VPN接続名";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "VPN ID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "状態";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "仮想プライベートゲートウェイ";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("F" + rowCounter);
            cell.Value2 = "カスタマーゲートウェイ";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var vpn in vpnList)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = getStringThroughTags(vpn.Tags);

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpn.VpnConnectionId;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = vpn.State.Value;

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = vpn.VpnGatewayId;

                cell = activeWorksheet.get_Range("F" + rowCounter);
                cell.Value2 = vpn.CustomerGatewayId;
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "F" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeVpn(List<VpnConnection> vpnList, int rowCounter)
        {
            foreach (var vpn in vpnList)
            {
                int instanceListStart = rowCounter + 2;

                // VPNエリア
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "■VPN";

                // タイトル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "項目名";
                cell.Interior.ColorIndex = colorIndex;

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "値";
                cell.Interior.ColorIndex = colorIndex;

                // VPN接続名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "VPN接続名";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = getStringThroughTags(vpn.Tags);

                // VPN ID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "VPN ID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpn.VpnConnectionId;

                // 状態
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "状態";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpn.State.Value;

                // 仮想プライベートゲートウェイ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "仮想プライベートゲートウェイ";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpn.VpnGatewayId;

                // カスタマーゲートウェイ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "カスタマーゲートウェイ";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpn.CustomerGatewayId;

                // タイプ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "タイプ";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vpn.Type.Value;

                // カテゴリ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "カテゴリ";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "VPN";

                // VPC
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "VPC";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                //cell.Value2 = "VPN";

                int instanceListEnd = rowCounter;

                // インスタンス一覧に罫線を引く
                activeWorksheet.get_Range("B" + instanceListStart + ":" + "C" + instanceListEnd).Borders.LineStyle = true;
                rowCounter++;
            }

            return rowCounter;
        }
      
        private void displayVgwButton_Click(object sender, RibbonControlEventArgs e)
        {
            #region 共通処理
            commonAction("VGW仕様書", false);
            #endregion

            #region ■VGW一覧の取得 
            var request = new DescribeVpnGatewaysRequest();

            List<VpnGateway> response = ec2Client.DescribeVpnGatewaysAsync(request).Result.VpnGateways;
            #endregion

            #region ■VGW一覧の描画
            rowCounter = makeVgwList(response, rowCounter);
            #endregion

            #region ■VGWの描画
            rowCounter = makeVgw(response, rowCounter);
            #endregion

            // 列幅を自動調整する
            activeWorksheet.get_Range("B1:H" + rowCounter).Columns.AutoFit();
        }

        private int makeVgwList(List<VpnGateway> vpnGateway, int rowCounter)
        {
            // 仮想プライベートゲートウェイエリア
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "■仮想プライベートゲートウェイ一覧";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "仮想プライベートゲートウェイ名";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "仮想プライベートゲートウェイID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "状態";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "タイプ";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("F" + rowCounter);
            cell.Value2 = "VPC";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("G" + rowCounter);
            cell.Value2 = "ASN (Amazon 側)";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var vgw in vpnGateway)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = getStringThroughTags(vgw.Tags);

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vgw.VpnGatewayId;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = vgw.State.Value;

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = vgw.Type.Value;

                cell = activeWorksheet.get_Range("F" + rowCounter);
                cell.Value2 = vgw.VpcAttachments[0].VpcId;

                cell = activeWorksheet.get_Range("G" + rowCounter);
                cell.Value2 = "64512";
                cell.NumberFormatLocal = "@";
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "G" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeVgw(List<VpnGateway> vpnGateway, int rowCounter)
        {
            foreach (var vgw in vpnGateway)
            {
                int instanceListStart = rowCounter + 2;

                // VGWエリア
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "■仮想プライベートゲートウェイ";

                // タイトル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "項目名";
                cell.Interior.ColorIndex = colorIndex;

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "値";
                cell.Interior.ColorIndex = colorIndex;

                // 仮想プライベートゲートウェイ名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "仮想プライベートゲートウェイ名";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = getStringThroughTags(vgw.Tags);

                // 仮想プライベートゲートウェイID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "仮想プライベートゲートウェイID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vgw.VpnGatewayId;

                // 状態
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "状態";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vgw.State.Value;

                // タイプ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "タイプ";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vgw.Type.Value;

                // VPC
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "VPC";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = vgw.VpcAttachments[0].VpcId;

                // ASN (Amazon 側)
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "ASN (Amazon 側)";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "64512";
                cell.NumberFormatLocal = "@";

                int instanceListEnd = rowCounter;

                // インスタンス一覧に罫線を引く
                activeWorksheet.get_Range("B" + instanceListStart + ":" + "C" + instanceListEnd).Borders.LineStyle = true;
                rowCounter++;
            }

            return rowCounter;
        }

        private void displayCgwButton_Click(object sender, RibbonControlEventArgs e)
        {
            #region 共通処理
            commonAction("CGW仕様書", false);
            #endregion

            #region ■CGW一覧の取得 
            var request = new DescribeCustomerGatewaysRequest();

            List<CustomerGateway> response = ec2Client.DescribeCustomerGatewaysAsync(request).Result.CustomerGateways;
            #endregion

            #region ■CGW一覧の描画
            rowCounter = makeCgwList(response, rowCounter);
            #endregion

            #region ■CGWの描画
            rowCounter = makeCgw(response, rowCounter);
            #endregion

            // 列幅を自動調整する
            activeWorksheet.get_Range("B1:H" + rowCounter).Columns.AutoFit();
        }

        private int makeCgwList(List<CustomerGateway> customerGateway, int rowCounter)
        {
            // カスタマーゲートウェイエリア
            rowCounter++;
            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "■カスタマーゲートウェイ一覧";

            // タイトル
            rowCounter++;
            int instanceListStart = rowCounter;

            cell = activeWorksheet.get_Range("B" + rowCounter);
            cell.Value2 = "カスタマーゲートウェイ名";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("C" + rowCounter);
            cell.Value2 = "カスタマーゲートウェイID";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("D" + rowCounter);
            cell.Value2 = "状態";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("E" + rowCounter);
            cell.Value2 = "タイプ";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("F" + rowCounter);
            cell.Value2 = "IP アドレス";
            cell.Interior.ColorIndex = colorIndex;

            cell = activeWorksheet.get_Range("G" + rowCounter);
            cell.Value2 = "BGP ASN";
            cell.Interior.ColorIndex = colorIndex;

            foreach (var cgw in customerGateway)
            {
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = getStringThroughTags(cgw.Tags);

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = cgw.CustomerGatewayId;

                cell = activeWorksheet.get_Range("D" + rowCounter);
                cell.Value2 = cgw.State;

                cell = activeWorksheet.get_Range("E" + rowCounter);
                cell.Value2 = cgw.Type;

                cell = activeWorksheet.get_Range("F" + rowCounter);
                cell.Value2 = cgw.IpAddress;

                cell = activeWorksheet.get_Range("G" + rowCounter);
                cell.Value2 = cgw.BgpAsn;
            }
            int instanceListEnd = rowCounter;

            // インスタンス一覧に罫線を引く
            activeWorksheet.get_Range("B" + instanceListStart + ":" + "G" + instanceListEnd).Borders.LineStyle = true;

            // 1行開ける
            rowCounter++;

            return rowCounter;
        }

        private int makeCgw(List<CustomerGateway> customerGateway, int rowCounter)
        {
            foreach (var cgw in customerGateway)
            {
                int instanceListStart = rowCounter + 2;

                // CGWエリア
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "■カスタマーゲートウェイ";

                // タイトル
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "項目名";
                cell.Interior.ColorIndex = colorIndex;

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = "値";
                cell.Interior.ColorIndex = colorIndex;

                // カスタマーゲートウェイ名
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "カスタマーゲートウェイ名";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = getStringThroughTags(cgw.Tags);

                // カスタマーゲートウェイID
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "カスタマーゲートウェイID";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = cgw.CustomerGatewayId;

                // 状態
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "状態";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = cgw.State;

                // タイプ
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "タイプ";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = cgw.Type;

                // IP アドレス
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "IP アドレス";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = cgw.IpAddress;

                // BGP ASN
                rowCounter++;
                cell = activeWorksheet.get_Range("B" + rowCounter);
                cell.Value2 = "BGP ASN";

                cell = activeWorksheet.get_Range("C" + rowCounter);
                cell.Value2 = cgw.BgpAsn;

                int instanceListEnd = rowCounter;

                // インスタンス一覧に罫線を引く
                activeWorksheet.get_Range("B" + instanceListStart + ":" + "C" + instanceListEnd).Borders.LineStyle = true;
                rowCounter++;
            }

            return rowCounter;
        }

    }
}
